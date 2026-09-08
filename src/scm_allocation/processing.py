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

from scm_allocation.extraction import OUT_OF_SCOPE_BATAM_SUBJECT, extract_email, is_out_of_scope_email
from scm_allocation.ingestion.category_writer import CategoryWriteResult, OutlookCategoryWriter
from scm_allocation.ingestion.inspector import get_sender_address, inspect_email_item
from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.normalization import normalize_text
from scm_allocation.processing_categories import add_category_text, category_parts


SCRAPED_CATEGORY = "SCM Bot - Scraped"
REVIEW_CATEGORY = "SCM Bot - Review"


@dataclass(frozen=True)
class CandidateFilter:
    """Conservative, configurable rules for identifying allocation emails."""

    sender_domains: tuple[str, ...] = ("erajaya.com",)
    keywords: tuple[str, ...] = ("alokasi", "allocation")
    excluded_subjects: tuple[str, ...] = (OUT_OF_SCOPE_BATAM_SUBJECT,)

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
        sender_domain = sender.rsplit("@", 1)[-1].casefold() if "@" in sender else ""
        if self.sender_domains and sender_domain not in {domain.casefold() for domain in self.sender_domains}:
            return False

        body = " ".join(
            value
            for value in (
                normalize_text(getattr(item, "Body", "")),
                normalize_text(getattr(item, "HTMLBody", "")),
            )
            if value
        ).casefold()
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
            item = items[index]
            self.last_scan_inspected += 1
            subject = normalize_text(getattr(item, "Subject", "")) or ""
            if self.config.candidate_filter.is_excluded(subject):
                self.last_invalid_or_excluded_count += 1
                continue
            if not self.config.candidate_filter.matches(item):
                continue
            self.last_scan_candidates += 1
            email_id = normalize_text(getattr(item, "EntryID", "")) or f"outlook-item-{index}"
            existing_categories = getattr(item, "Categories", "") or ""
            if {
                SCRAPED_CATEGORY.casefold(), REVIEW_CATEGORY.casefold()
            } & {
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
                self.last_error_count += bool(summary.error_information)
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
                + (f" | failure: {summary.metadata.get('category_write_reason')}" if not summary.category_written and summary.category else "")
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
            folder = folders[index]
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
            metadata=result.metadata,
        )
        self.repository.record_result(email_id, summary)
        return summary

    def _extract_inspected_email(self, item: Any, inspected: Mapping[str, object]) -> ExtractionResult:
        excel_attachments = [
            attachment
            for attachment in getattr(item, "Attachments", [])
            if str(getattr(attachment, "FileName", "")).casefold().endswith(".xlsx")
        ]
        if not excel_attachments:
            return self.extractor(inspected)

        with TemporaryDirectory(prefix="scm-allocation-") as temp_dir:
            attachment_paths: dict[str, Path] = {}
            for attachment in excel_attachments:
                filename = str(getattr(attachment, "FileName", ""))
                safe_filename = Path(filename).name
                target = Path(temp_dir) / safe_filename
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
