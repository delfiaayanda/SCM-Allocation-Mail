"""Mocked tests for Outlook candidate processing and category state."""

from __future__ import annotations

from pathlib import Path

import pytest

from scm_allocation.models.allocation import ExtractionResult, ExtractionStatus
from scm_allocation.ingestion.category_writer import OutlookCategoryWriter
from scm_allocation.processing import (
    REVIEW_CATEGORY,
    SCRAPED_CATEGORY,
    CandidateFilter,
    InMemoryEmailProcessingRepository,
    OutlookEmailScanner,
    OutlookScanConfig,
    ProcessingSummary,
    add_category,
)


class FakeAttachment:
    def __init__(self, filename: str = "allocation.xlsx") -> None:
        self.FileName = filename
        self.Size = 10
        self.saved_paths: list[str] = []

    def SaveAsFile(self, path: str) -> None:
        self.saved_paths.append(path)
        Path(path).write_bytes(b"synthetic")


class FakeItem:
    def __init__(
        self,
        *,
        subject: str = "Alokasi sample request",
        categories: str = "Blue Category, Important",
        attachment: FakeAttachment | None = None,
    ) -> None:
        self.EntryID = "stable-entry-id"
        self.Subject = subject
        self.SenderEmailType = "SMTP"
        self.SenderEmailAddress = "pm@erajaya.com"
        self.SenderName = "PM"
        self.ReceivedTime = "2026-09-07 15:00:00+00:00"
        self.Body = "Mohon bantu alokasi sample."
        self.HTMLBody = "<html><body>Mohon bantu alokasi sample.</body></html>"
        self.UnRead = True
        self.Categories = categories
        self.Attachments = [attachment] if attachment else []
        self.Application = None
        self.save_calls = 0

    def Save(self):
        self.save_calls += 1


class FakeCategory:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeCategories:
    def __init__(self, names=()):
        self.values = [FakeCategory(name) for name in names]
        self.Add_calls = []

    @property
    def Count(self):
        return len(self.values)

    def __getitem__(self, index):
        return self.values[index - 1]

    def Add(self, name):
        self.Add_calls.append(name)
        category = FakeCategory(name)
        self.values.append(category)
        return category


class FakeItems:
    def __init__(self, item: FakeItem) -> None:
        self._item = item
        self.Count = 1
        self.sorted = False

    def Sort(self, field: str, descending: bool) -> None:
        self.sorted = True

    def __getitem__(self, index: int) -> FakeItem:
        assert index == 1
        return self._item


class FakeFolder:
    def __init__(self, item: FakeItem) -> None:
        self.Items = FakeItems(item)


class FakeNamespace:
    def __init__(self, folder: FakeFolder) -> None:
        self.folder = folder
        self.Categories = FakeCategories()

    def GetDefaultFolder(self, folder_id: int) -> FakeFolder:
        assert folder_id == 6
        return self.folder


class FakeOutlook:
    def __init__(self, folder: FakeFolder) -> None:
        self.namespace = FakeNamespace(folder)

    def GetNamespace(self, name: str) -> FakeNamespace:
        assert name == "MAPI"
        return self.namespace


def result(status: ExtractionStatus) -> ExtractionResult:
    return ExtractionResult([], status, "test_parser", "stable-entry-id", ("test detail",))


def scanner_for(item: FakeItem, extraction: ExtractionResult, *, dry_run: bool, repository=None, calls=None):
    def extractor(email, **kwargs):
        if calls is not None:
            calls.append(email["entry_id"])
        return extraction

    return OutlookEmailScanner(
        repository=repository,
        config=OutlookScanConfig(dry_run=dry_run, max_emails=1),
        extractor=extractor,
    ), FakeOutlook(FakeFolder(item))


def test_dry_run_does_not_write_categories_or_read_state():
    item = FakeItem(categories="Blue Category")
    scanner, outlook = scanner_for(item, result(ExtractionStatus.VALID), dry_run=True)

    summaries = scanner.scan(outlook)

    assert summaries[0].category == SCRAPED_CATEGORY
    assert summaries[0].category_written is False
    assert item.Categories == "Blue Category"
    assert item.UnRead is True
    assert item.save_calls == 0
    assert summaries[0].unread_status_preserved is True


@pytest.mark.parametrize(
    ("status", "expected_category"),
    [
        (ExtractionStatus.VALID, SCRAPED_CATEGORY),
        (ExtractionStatus.PARTIAL, REVIEW_CATEGORY),
        (ExtractionStatus.INVALID, None),
        (ExtractionStatus.UNSUPPORTED, None),
    ],
)
def test_write_mode_assigns_category_by_extraction_status(status, expected_category):
    item = FakeItem(categories="Blue Category, Important")
    scanner, outlook = scanner_for(item, result(status), dry_run=False)

    summaries = scanner.scan(outlook)

    assert summaries[0].category == expected_category
    assert summaries[0].category_written is (expected_category is not None)
    if expected_category:
        assert item.Categories == f"Blue Category, Important, {expected_category}"
    else:
        assert item.Categories == "Blue Category, Important"
    assert item.UnRead is True
    assert item.save_calls == (1 if expected_category else 0)


def test_existing_bot_category_is_preserved_without_duplicate():
    item = FakeItem(categories=f"Blue Category, {SCRAPED_CATEGORY}")
    scanner, outlook = scanner_for(item, result(ExtractionStatus.VALID), dry_run=False)

    summary = scanner.scan(outlook)[0]

    assert summary.skipped is True
    assert summary.category_written is False
    assert item.Categories == f"Blue Category, {SCRAPED_CATEGORY}"


def test_repository_prevents_duplicate_valid_processing():
    item = FakeItem(categories="")
    calls: list[str] = []
    repository = InMemoryEmailProcessingRepository()
    scanner, outlook = scanner_for(item, result(ExtractionStatus.VALID), dry_run=True, repository=repository, calls=calls)

    first = scanner.scan(outlook)[0]
    second = scanner.scan(outlook)[0]

    assert first.skipped is False
    assert second.skipped is True
    assert calls == ["stable-entry-id"]
    assert repository.is_processed("stable-entry-id") is True
    assert repository.get_processing_status("stable-entry-id") == first


def test_batam_email_is_not_a_candidate_and_is_not_modified():
    item = FakeItem(subject="RE: STO DEVICE ALOKASI BATAM Asia Brand WK36 2026", categories="Blue Category")
    calls: list[str] = []
    scanner, outlook = scanner_for(item, result(ExtractionStatus.VALID), dry_run=False, calls=calls)

    summaries = scanner.scan(outlook)

    assert summaries == []
    assert calls == []
    assert item.Categories == "Blue Category"
    assert item.UnRead is True


def test_candidate_filter_is_conservative_and_configurable():
    non_candidate = FakeItem(subject="Weekly team announcement")
    non_candidate.Body = "Routine status update."
    non_candidate.HTMLBody = "<html><body>Routine status update.</body></html>"
    assert CandidateFilter().matches(non_candidate) is False
    assert CandidateFilter(sender_domains=(), keywords=("announcement",)).matches(non_candidate) is True


def test_add_category_handles_empty_and_existing_delimiters():
    assert add_category("", SCRAPED_CATEGORY) == SCRAPED_CATEGORY
    assert add_category("Blue; Important", REVIEW_CATEGORY) == "Blue, Important, SCM Bot - Review"
    assert add_category(f"Blue, {SCRAPED_CATEGORY}", SCRAPED_CATEGORY) == f"Blue, {SCRAPED_CATEGORY}"


def test_scanner_reports_connection_folder_and_scan_counts():
    item = FakeItem()
    messages: list[str] = []
    scanner, outlook = scanner_for(item, result(ExtractionStatus.VALID), dry_run=True)
    scanner.progress_callback = messages.append

    scanner.scan(outlook)

    assert messages == [
        "Connected to Outlook.",
        "Scanning folder: Inbox.",
        "Outlook folder contains 1 item(s); inspecting 1 newest item(s).",
        "Inspected 1 item(s); found 1 candidate(s).",
    ]
    assert scanner.last_scan_inspected == 1
    assert scanner.last_scan_candidates == 1


def test_processing_summary_keeps_error_details():
    item = FakeItem()
    scanner, outlook = scanner_for(item, result(ExtractionStatus.INVALID), dry_run=True)

    summary = scanner.scan(outlook)[0]

    assert summary.error_information == ("test detail",)
    assert summary.category is None
    assert summary.category_written is False


def test_category_writer_reuses_existing_and_is_idempotent():
    categories = FakeCategories([SCRAPED_CATEGORY])
    item = FakeItem(categories="Blue Category")
    writer = OutlookCategoryWriter(categories)

    first = writer.apply_category(item, SCRAPED_CATEGORY)
    second = writer.apply_category(item, SCRAPED_CATEGORY)

    assert first.success is True
    assert first.written is True
    assert second.success is True
    assert second.written is False
    assert categories.Add_calls == []
    assert item.UnRead is True


def test_category_writer_creates_missing_category_once():
    categories = FakeCategories()
    item = FakeItem(categories="Blue Category")
    writer = OutlookCategoryWriter(categories)

    outcome = writer.apply_category(item, REVIEW_CATEGORY)

    assert outcome.success is True
    assert outcome.written is True
    assert categories.Add_calls == [REVIEW_CATEGORY]
    assert item.Categories == f"Blue Category, {REVIEW_CATEGORY}"


def test_category_writer_reports_failure_without_read_state_change():
    class FailingCategories(FakeCategories):
        def Add(self, name):
            raise RuntimeError("category service unavailable")

    categories = FailingCategories()
    item = FakeItem(categories="Blue Category")
    item.Save = lambda: (_ for _ in ()).throw(RuntimeError("save service unavailable"))
    before = item.UnRead
    outcome = OutlookCategoryWriter(categories).apply_category(item, SCRAPED_CATEGORY)

    assert outcome.success is False
    assert outcome.written is False
    assert "unavailable" in outcome.reason
    assert item.Categories == "Blue Category"
    assert item.UnRead == before


def test_category_writer_falls_back_when_category_collection_rejects_add():
    class RejectingCategories(FakeCategories):
        def Add(self, *args):
            raise RuntimeError("Categories.Add rejected")

    item = FakeItem(categories="Blue Category")
    outcome = OutlookCategoryWriter(RejectingCategories()).apply_category(item, REVIEW_CATEGORY)

    assert outcome.success is True
    assert outcome.written is True
    assert "rejected" in outcome.reason
    assert item.Categories == f"Blue Category, {REVIEW_CATEGORY}"


def test_category_failure_does_not_stop_later_category_write():
    class FailingItem(FakeItem):
        def __setattr__(self, name, value):
            if name == "Categories" and hasattr(self, "Categories"):
                raise RuntimeError("first item cannot accept category")
            super().__setattr__(name, value)

    first = FailingItem(categories="Blue Category")
    second = FakeItem(categories="Blue Category")
    categories = FakeCategories()
    scanner = OutlookEmailScanner(config=OutlookScanConfig(dry_run=False))
    summaries = [
        ProcessingSummary("one", "One", "one@example.com", "now", "now", ExtractionStatus.VALID, "test", 1, SCRAPED_CATEGORY, False),
        ProcessingSummary("two", "Two", "two@example.com", "now", "now", ExtractionStatus.PARTIAL, "test", 1, REVIEW_CATEGORY, False),
    ]

    scanner._write_categories([(first, summaries[0]), (second, summaries[1])], categories)

    assert summaries[0].category_written is False
    assert "cannot accept" in summaries[0].metadata["category_write_reason"]
    assert summaries[1].category_written is True
    assert second.Categories == f"Blue Category, {REVIEW_CATEGORY}"