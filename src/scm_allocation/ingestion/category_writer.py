"""Isolated Outlook category creation and assignment adapter."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from scm_allocation.processing_categories import add_category_text, category_parts


@dataclass(frozen=True)
class CategoryWriteResult:
    """Outcome of one idempotent category assignment."""

    success: bool
    written: bool
    category: str
    reason: str = ""


class OutlookCategoryWriter:
    """Write only Outlook Categories through an injected Outlook object model."""

    def __init__(self, categories: Any) -> None:
        self.categories = categories

    def ensure_category(self, name: str) -> Any:
        """Reuse an existing category or create it once."""
        for index in range(1, int(getattr(self.categories, "Count", 0)) + 1):
            try:
                category = self.categories[index]
            except (IndexError, KeyError):
                break
            if str(getattr(category, "Name", "")).casefold() == name.casefold():
                return category
        try:
            return self.categories.Add(name, 0)
        except TypeError:
            return self.categories.Add(name)

    def apply_category(self, item: Any, category_name: str) -> CategoryWriteResult:
        """Apply a category without changing any other MailItem property."""
        existing = str(getattr(item, "Categories", "") or "")
        try:
            if category_name.casefold() in {value.casefold() for value in category_parts(existing)}:
                return CategoryWriteResult(True, False, category_name, "category already present")
            category_creation_reason = ""
            try:
                self.ensure_category(category_name)
            except Exception as err:
                # Some Outlook profiles reject Categories.Add while accepting
                # the MailItem Categories property itself.
                category_creation_reason = f"category object reuse/create unavailable: {err}"
            item.Categories = add_category_text(existing, category_name)
            save = getattr(item, "Save", None)
            if callable(save):
                save()
            return CategoryWriteResult(True, True, category_name, category_creation_reason)
        except Exception as err:
            try:
                item.Categories = existing
            except Exception:
                pass
            return CategoryWriteResult(False, False, category_name, str(err))
