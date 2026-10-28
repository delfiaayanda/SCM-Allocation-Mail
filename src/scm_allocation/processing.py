"""Read-only Outlook scanning with explicit processing-category writes."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from typing import Any, Callable, Dict, Mapping, Optional

from scm_allocation.allocation_subject import NON_ALLOCATION_SUBJECT_RE, is_allocation_subject, non_request_reason
from scm_allocation.extraction import OUT_OF_SCOPE_BATAM_SUBJECT, extract_email, is_allocation_attachment_filename, is_out_of_scope_email
from scm_allocation.ingestion.category_writer import CategoryWriteResult, OutlookCategoryWriter
from scm_allocation.ingestion.inspector import get_sender_address, inspect_email_item
from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.normalization import normalize_text
from scm_allocation.processing_categories import add_category_text, category_parts


SCRAPED_CATEGORY = "SCM Bot - Scraped"
REVIEW_CATEGORY = "SCM Bot - Review"
# Email dari pengirim di daftar wajib APEX (sheet MNP/CM -> PIC RICO) yang tidak
# berisi alokasi terbaca: tetap di-push ke APEX sebagai email info (isi ringkas).
INFO_CATEGORY = "SCM Bot - Info"
# Email ALOKASI (dari subjek, lihat allocation_subject.py) yang isinya tidak bisa
# diparse: tetap di-push ke APEX sebagai jenis 'tak_terbaca' untuk dikerjakan
# manual, lalu diberi kategori ini supaya tidak diproses ulang tiap jam.
UNREADABLE_CATEGORY = "SCM Bot - Tak Terbaca"
BOT_CATEGORIES = (SCRAPED_CATEGORY, REVIEW_CATEGORY, INFO_CATEGORY, UNREADABLE_CATEGORY)
BODY_PREVIEW_CHARS = 2000


@dataclass(frozen=True)
class CandidateFilter:
    """Conservative, configurable rules for identifying allocation emails."""

    sender_domains: tuple[str, ...] = ("erajaya.com",)
    keywords: tuple[str, ...] = ("alokasi", "allocation")
    excluded_subjects: tuple[str, ...] = (OUT_OF_SCOPE_BATAM_SUBJECT,)
    # Pengirim (huruf kecil) yang emailnya WAJIB diambil walau tidak berbau
    # alokasi dan walau domainnya bukan sender_domains — daftar dari APEX
    # (/api/submissions/pengirim-wajib), lihat scripts/process_outlook_allocations.py.
    wajib_senders: frozenset[str] = frozenset()

    def is_wajib(self, sender: object) -> bool:
        return bool(self.wajib_senders) and str(sender or "").strip().casefold() in self.wajib_senders

    def is_excluded(self, subject: object) -> bool:
        """Return whether an email is explicitly outside the allocation scope."""
        normalized_subject = normalize_text(subject) or ""
        return (
            normalized_subject.casefold() in {value.casefold() for value in self.excluded_subjects}
            or is_out_of_scope_email(normalized_subject)
        )

    def matches(self, item: Any) -> bool:
        subject = normalize_text(getattr(item, "Subject", "")) or ""
        if self.is_excluded(subject):
            return False

        sender = get_sender_address(item)
        if self.is_wajib(sender):
            return True
        # Subjek alokasi selalu kandidat, dari domain mana pun (akun Gmail toko,
        # subjek "REQUEST STOCK … NEW STORE" tanpa kata alokasi).
        if is_allocation_subject(subject, sender):
            return True
        # Retur / cancel STO / notifikasi: dulu tetap lolos lewat kata "alokasi"
        # di isi email dan terbaca sebagai alokasi (Request Returan Terpusat:
        # 34 rb baris) — padahal bukan permintaan alokasi.
        if NON_ALLOCATION_SUBJECT_RE.search(subject):
            return False
        sender_domain = sender.rsplit("@", 1)[-1].casefold() if "@" in sender else ""
        if self.sender_domains and sender_domain not in {domain.casefold() for domain in self.sender_domains}:
            return False

        try:
            body = " ".join(
                value
                for value in (
                    normalize_text(getattr(item, "Body", "")),
                    normalize_text(getattr(item, "HTMLBody", "")),
                )
                if value
            ).casefold()
        except Exception:
            # Some Outlook item types (meeting requests, read receipts, sharing
            # invites) raise a COM "Not implemented" error reading Body/HTMLBody
            # instead of returning "" — getattr's default only covers
            # AttributeError, not COM errors, so this must be caught explicitly
            # or one such item aborts the entire folder scan.
            body = ""
        searchable = f"{subject.casefold()} {body}"
        return any(keyword.casefold() in searchable for keyword in self.keywords)


@dataclass(frozen=True)
class OutlookScanConfig:
    """Runtime settings for one bounded Outlook scan."""

    folder_name: str = "Inbox"
    max_emails: Optional[int] = None
    dry_run: bool = True
    candidate_filter: CandidateFilter = field(default_factory=CandidateFilter)


@dataclass
class ProcessingSummary:
    """Auditable result for one candidate email."""

    email_id: str
    subject: str
    sender: str
    received_at: str
    processing_timestamp: str
    extraction_status: ExtractionStatus
    parser_type: str
    record_count: int
    category: Optional[str]
    category_written: bool
    error_information: tuple[str, ...] = ()
    review_information: tuple[str, ...] = ()
    unread_status_preserved: bool = True
    skipped: bool = False
    metadata: Mapping[str, object] = field(default_factory=dict)


class EmailProcessingRepository(ABC):
    """Persistence boundary for email processing summaries."""

    @abstractmethod
    def is_processed(self, email_id: str) -> bool:
        """Return whether a valid result was already recorded."""
        raise NotImplementedError

    @abstractmethod
    def record_result(self, email_id: str, result: ProcessingSummary) -> None:
        """Record the latest processing summary for an email."""
        raise NotImplementedError

    @abstractmethod
    def get_processing_status(self, email_id: str) -> Optional[ProcessingSummary]:
        """Return the recorded summary for an email, if any."""
        raise NotImplementedError


class InMemoryEmailProcessingRepository(EmailProcessingRepository):
    """Local development repository; replaceable by a future database adapter."""

    def __init__(self) -> None:
        self._results: Dict[str, ProcessingSummary] = {}

    def is_processed(self, email_id: str) -> bool:
        result = self._results.get(email_id)
        return result is not None and result.extraction_status is ExtractionStatus.VALID

    def record_result(self, email_id: str, result: ProcessingSummary) -> None:
        self._results[email_id] = result

    def get_processing_status(self, email_id: str) -> Optional[ProcessingSummary]:
        return self._results.get(email_id)


def _category_parts(categories: object) -> list[str]:
    return category_parts(categories)


def add_category(categories: object, category: str) -> str:
    """Return categories with one new value while preserving existing values."""
    return add_category_text(categories, category)


def category_for_result(result: ExtractionResult) -> Optional[str]:
    if result.status is ExtractionStatus.VALID:
        return SCRAPED_CATEGORY
    if result.status is ExtractionStatus.PARTIAL:
        return REVIEW_CATEGORY
    return None


def _body_preview(item: Any) -> str:
    """Isi teks email, dipotong BODY_PREVIEW_CHARS. Item COM tertentu melempar
    error saat Body dibaca (lihat CandidateFilter.matches) — jadi kosong saja."""
    try:
        body = str(getattr(item, "Body", "") or "")
    except Exception:
        body = ""
    # Baris dipertahankan (normalize_text meratakan semua whitespace jadi satu
    # spasi); cuma spasi berlebih per baris dan baris kosong beruntun yang dirapikan.
    baris = [re.sub(r"[ \t\u00a0]+", " ", b).strip() for b in body.replace("\r\n", "\n").split("\n")]
    body = re.sub(r"\n{3,}", "\n\n", "\n".join(baris)).strip()
    return body[:BODY_PREVIEW_CHARS]


def _unreadable_reason(result: ExtractionResult, inspected: Optional[Mapping[str, object]]) -> str:
    """Kalimat untuk PIC di APEX: kenapa bot gagal + lampiran apa saja yang ada."""
    reasons = {
        "no supported allocation table found": "isi email tidak berisi tabel alokasi yang dikenali",
        "no allocation worksheet found": "lampiran Excel tidak berisi sheet alokasi yang dikenali",
        "allocation worksheet(s) found but no records extracted": "sheet alokasi ada tapi qty-nya kosong semua",
        "email has no supported Excel attachment or HTML body": "tidak ada lampiran Excel maupun tabel di isi email",
    }
    def explain(error: str) -> str:
        if error in reasons:
            return reasons[error]
        if "unavailable locally" in error or "zip file" in error.casefold() or "not a zip" in error.casefold():
            return "lampiran tidak bisa dibuka (rusak, terkunci sandi, atau tautan OneDrive/SharePoint)"
        return error
    detail = "; ".join(dict.fromkeys(explain(error) for error in result.errors)) or "tidak ada alokasi terbaca"
    names = [
        str(attachment.get("filename"))
        for attachment in ((inspected or {}).get("attachments") or [])
        if isinstance(attachment, Mapping) and attachment.get("filename")
        and not re.fullmatch(r"(?i)image\d+\.(png|jpe?g|gif)|untitled attachment.*\.txt|.*\.(ics)", str(attachment.get("filename")))
    ]
    return f"Bot: {detail}." + (f" Lampiran: {', '.join(names)}." if names else " Tanpa lampiran.")


def _error_result(email_id: str, parser_type: str, message: str) -> ExtractionResult:
    return ExtractionResult([], ExtractionStatus.INVALID, parser_type, email_id, (message,))


class OutlookEmailScanner:
    """Scan a bounded Outlook folder and optionally write processing categories."""

    def __init__(
        self,
        repository: Optional[EmailProcessingRepository] = None,
        config: Optional[OutlookScanConfig] = None,
        extractor: Callable[..., ExtractionResult] = extract_email,
        progress_callback: Optional[Callable[[str], None]] = None,
        category_writer_factory: Optional[Callable[[Any], OutlookCategoryWriter]] = None,
        history_logger: Optional[Any] = None,
    ) -> None:
        self.repository = repository or InMemoryEmailProcessingRepository()
        self.config = config or OutlookScanConfig()
        self.extractor = extractor
        self.progress_callback = progress_callback
        self.category_writer_factory = category_writer_factory or OutlookCategoryWriter
        self.history_logger = history_logger
        self.last_scan_inspected = 0
        self.last_scan_candidates = 0
        self.last_scan_total = 0
        self.last_processed_count = 0
        self.last_skipped_already_processed = 0
        self.last_scraped_count = 0
        self.last_review_count = 0
        self.last_invalid_or_excluded_count = 0
        self.last_write_failure_count = 0
        self.last_error_count = 0

    def _report_progress(self, message: str) -> None:
        if self.progress_callback:
            self.progress_callback(message)

    def scan(self, outlook_application: Any = None) -> list[ProcessingSummary]:
        """Scan Outlook without changing categories unless dry_run is disabled."""
        outlook = outlook_application or self._connect_outlook()
        self._report_progress("Connected to Outlook.")
        namespace = outlook.GetNamespace("MAPI")
        folder = self._resolve_folder(namespace)
        self._report_progress(f"Scanning folder: {self.config.folder_name}.")
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        try:
            total_count = int(getattr(items, "Count", 0))
        except (TypeError, ValueError):
            total_count = 0
        fetch_count = total_count if self.config.max_emails is None else min(self.config.max_emails, total_count)
        self.last_scan_total = fetch_count
        self.last_scan_inspected = 0
        self.last_scan_candidates = 0
        self._report_progress(
            f"Outlook folder contains {total_count} item(s); inspecting {fetch_count} newest item(s)."
        )
        summaries: list[ProcessingSummary] = []
        pending_writes: list[tuple[Any, ProcessingSummary]] = []
        processed_items: list[tuple[Any, ProcessingSummary]] = []
        self.last_processed_count = 0
        self.last_skipped_already_processed = 0
        self.last_scraped_count = 0
        self.last_review_count = 0
        self.last_invalid_or_excluded_count = 0
        self.last_write_failure_count = 0
        self.last_error_count = 0

        for index in range(1, fetch_count + 1):
            self.last_scan_inspected += 1
            try:
                # .Item() = COM 1-based, sesuai range() di atas. JANGAN items[index]:
                # __getitem__ pywin32 0-based, jadi email terbaru (posisi 0)
                # terlewat dan index terakhir meledak "list index out of range".
                item = items.Item(index)
            except Exception as err:
                self.last_error_count += 1
                self._report_progress(f"ERROR retrieving Outlook item {index}: {err}")
                continue
            subject = normalize_text(getattr(item, "Subject", "")) or ""
            if self.config.candidate_filter.is_excluded(subject):
                self.last_invalid_or_excluded_count += 1
                continue
            if not self.config.candidate_filter.matches(item):
                continue
            self.last_scan_candidates += 1
            email_id = normalize_text(getattr(item, "EntryID", "")) or f"outlook-item-{index}"
            existing_categories = getattr(item, "Categories", "") or ""
            if {category.casefold() for category in BOT_CATEGORIES} & {
                part.casefold() for part in _category_parts(existing_categories)
            }:
                summaries.append(self._skipped_summary(item, email_id))
                self.last_skipped_already_processed += 1
                self._report_progress(f"SKIPPED_ALREADY_PROCESSED: {getattr(item, 'Subject', '<No Subject>')}")
                continue
            summary = self._process_item(item, index, email_id)
            summaries.append(summary)
            processed_items.append((item, summary))
            self.last_processed_count += 1
            if summary.extraction_status is ExtractionStatus.VALID:
                self.last_scraped_count += 1
            elif summary.extraction_status is ExtractionStatus.PARTIAL:
                self.last_review_count += 1
            else:
                self.last_invalid_or_excluded_count += 1
                self.last_error_count += 0
            if summary.category:
                pending_writes.append((item, summary))

        self._report_progress(
            f"Inspected {self.last_scan_inspected} item(s); found {self.last_scan_candidates} candidate(s)."
        )
        if not self.config.dry_run and pending_writes:
            self._report_write_summary(summaries, pending_writes)
            self._write_categories(pending_writes, getattr(namespace, "Categories", None))
        if self.history_logger and not self.config.dry_run:
            for item, summary in processed_items:
                if (
                    summary.extraction_status not in {ExtractionStatus.VALID, ExtractionStatus.PARTIAL}
                    or not summary.category_written
                ):
                    continue
                try:
                    attachment_count = int(getattr(getattr(item, "Attachments", []), "Count", 0))
                except (TypeError, ValueError):
                    attachment_count = 0
                self.history_logger.record_successful_processing(summary, attachment_count)
        return summaries

    def run_counters(self) -> dict[str, int]:
        return {
            "scanned_count": self.last_scan_inspected,
            "candidate_count": self.last_scan_candidates,
            "processed_count": self.last_processed_count,
            "skipped_already_processed": self.last_skipped_already_processed,
            "scraped_count": self.last_scraped_count,
            "review_count": self.last_review_count,
            "invalid_or_excluded_count": self.last_invalid_or_excluded_count,
            "write_failure_count": self.last_write_failure_count,
            "error_count": self.last_error_count,
        }

    def _report_write_summary(self, summaries: list[ProcessingSummary], pending_writes: list[tuple[Any, ProcessingSummary]]) -> None:
        valid_count = sum(summary.extraction_status is ExtractionStatus.VALID for summary in summaries)
        review_count = sum(summary.extraction_status is ExtractionStatus.PARTIAL for summary in summaries)
        self._report_progress(f"VALID emails: {valid_count}")
        self._report_progress(f"REVIEW/PARTIAL emails: {review_count}")
        self._report_progress(f"Emails receiving {SCRAPED_CATEGORY}: {sum(summary.category == SCRAPED_CATEGORY for _, summary in pending_writes)}")
        self._report_progress(f"Emails receiving SCM Bot - Review: {sum(summary.category == 'SCM Bot - Review' for _, summary in pending_writes)}")
        self._report_progress(f"INVALID/unsupported/excluded emails: {sum(summary.category is None for summary in summaries)}")
        self._report_progress(f"Total candidates: {len(summaries)}")
        self._report_progress("LIVE CATEGORY WRITE ENABLED")
        self._report_progress(f"The following {len(pending_writes)} emails will receive Outlook categories.")

    def _write_categories(self, pending_writes: list[tuple[Any, ProcessingSummary]], categories: Any) -> None:
        writers: dict[int, OutlookCategoryWriter] = {}
        for item, summary in pending_writes:
            try:
                category_collection = categories
                if category_collection is None:
                    category_collection = item.Application.Session.Categories
                writer = writers.setdefault(id(category_collection), self.category_writer_factory(category_collection))
                outcome = writer.apply_category(item, summary.category or "")
                summary.category_written = outcome.written
                summary.metadata["category_write_reason"] = outcome.reason
                if not outcome.written and outcome.reason != "category already present":
                    self.last_write_failure_count += 1
            except Exception as err:
                summary.category_written = False
                summary.metadata["category_write_reason"] = str(err)
                self.last_write_failure_count += 1
            self._report_progress(
                f"Subject: {summary.subject} | Status: {summary.extraction_status.value.upper()} | "
                f"Planned category: {summary.category or 'None'} | Written: {summary.category_written}"
                + f" | Write reason: {summary.metadata.get('category_write_reason') or 'written'}"
            )

    def _connect_outlook(self) -> Any:
        if sys.platform != "win32":
            raise RuntimeError("Outlook COM automation is only supported on Windows.")
        try:
            import win32com.client
        except ImportError as err:
            raise RuntimeError("pywin32 is required for Outlook automation.") from err
        try:
            return win32com.client.Dispatch("Outlook.Application")
        except Exception as err:
            raise RuntimeError(f"Failed to connect to Outlook desktop: {err}") from err

    def _resolve_folder(self, namespace: Any) -> Any:
        if self.config.folder_name.casefold() == "inbox":
            return namespace.GetDefaultFolder(6)
        folders = namespace.Folders
        for index in range(1, int(getattr(folders, "Count", 0)) + 1):
            folder = folders.Item(index)  # 1-based COM, bukan __getitem__ 0-based pywin32
            if normalize_text(getattr(folder, "Name", "")).casefold() == self.config.folder_name.casefold():
                return folder
        raise RuntimeError(f"Outlook folder not found: {self.config.folder_name}")

    def _process_item(self, item: Any, index: int, email_id: str) -> ProcessingSummary:
        unread_before = getattr(item, "UnRead", None)
        subject = normalize_text(getattr(item, "Subject", "")) or "<No Subject>"
        sender = get_sender_address(item)
        received_at = str(getattr(item, "ReceivedTime", "<Unknown>"))
        inspected = None
        try:
            inspected = inspect_email_item(item, index)
            result = self._extract_inspected_email(item, inspected)
        except Exception as err:
            result = _error_result(email_id, "outlook_processing", str(err))

        desired_category = category_for_result(result)
        info_metadata: dict[str, object] = {}
        if (
            result.status in {ExtractionStatus.INVALID, ExtractionStatus.UNSUPPORTED}
            and not result.records
            and is_allocation_subject(subject, sender)
        ):
            # Email alokasi yang isinya tidak terbaca: JANGAN dibuang diam-diam —
            # kirim ke APEX dengan alasannya supaya PIC mengerjakan manual.
            # Kecuali balasan thread tanpa permintaan baru / memo: masuk sebagai
            # info supaya tidak mengantre jadi kerja manual.
            has_excel = any(
                isinstance(a, Mapping) and is_allocation_attachment_filename(a.get("filename"))
                for a in ((inspected or {}).get("attachments") or [])
            )
            catatan = non_request_reason(subject, (inspected or {}).get("body_plain_text"), has_excel)
            if catatan:
                desired_category = INFO_CATEGORY
                info_metadata = {"catatan_bot": catatan, "body_preview": _body_preview(item)}
            else:
                desired_category = UNREADABLE_CATEGORY
                info_metadata = {
                    "tak_terbaca": _unreadable_reason(result, inspected),
                    "body_preview": _body_preview(item),
                }
        elif not result.records and self.config.candidate_filter.is_wajib(sender):
            # Pengirim di daftar wajib, tapi tidak ada alokasi terbaca: tetap
            # dikirim ke APEX sebagai email info supaya masuk kotak masuk PIC-nya.
            desired_category = INFO_CATEGORY
            info_metadata = {"wajib": True, "body_preview": _body_preview(item)}

        unread_after = getattr(item, "UnRead", None)
        errors = result.errors if result.status in {ExtractionStatus.INVALID, ExtractionStatus.UNSUPPORTED} else ()
        review = result.errors if result.status is ExtractionStatus.PARTIAL else ()
        summary = ProcessingSummary(
            email_id=email_id,
            subject=subject,
            sender=sender,
            received_at=received_at,
            processing_timestamp=datetime.now(timezone.utc).isoformat(),
            extraction_status=result.status,
            parser_type=result.parser_type,
            record_count=len(result.records),
            category=desired_category,
            category_written=False,
            error_information=errors,
            review_information=review,
            unread_status_preserved=unread_before == unread_after,
            metadata={**result.metadata, "records": result.records, **info_metadata},
        )
        self.repository.record_result(email_id, summary)
        return summary

    def _extract_inspected_email(self, item: Any, inspected: Mapping[str, object]) -> ExtractionResult:
        excel_attachments = [
            attachment
            for attachment in getattr(item, "Attachments", [])
            if is_allocation_attachment_filename(getattr(attachment, "FileName", ""))
        ]
        if not excel_attachments:
            return self.extractor(inspected)

        with TemporaryDirectory(prefix="scm-allocation-") as temp_dir:
            attachment_paths: dict[str, Path] = {}
            for number, attachment in enumerate(excel_attachments):
                filename = str(getattr(attachment, "FileName", ""))
                safe_filename = Path(filename).name
                # Satu subfolder per lampiran: dua lampiran bernama sama tidak saling timpa.
                target = Path(temp_dir) / str(number) / safe_filename
                target.parent.mkdir()
                attachment.SaveAsFile(str(target))
                attachment_paths[filename] = target
            return self.extractor(inspected, attachment_paths=attachment_paths)

    def _skipped_summary(self, item: Any, email_id: str) -> ProcessingSummary:
        return ProcessingSummary(
            email_id=email_id,
            subject=normalize_text(getattr(item, "Subject", "")) or "<No Subject>",
            sender=get_sender_address(item),
            received_at=str(getattr(item, "ReceivedTime", "<Unknown>")),
            processing_timestamp=datetime.now(timezone.utc).isoformat(),
            extraction_status=ExtractionStatus.SKIPPED,
            parser_type="duplicate_guard",
            record_count=0,
            category=None,
            category_written=False,
            skipped=True,
        )
