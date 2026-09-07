"""Shared Outlook category text helpers."""

from __future__ import annotations

import re

from scm_allocation.normalization import normalize_text


def category_parts(categories: object) -> list[str]:
    text = normalize_text(categories)
    if not text:
        return []
    return [part.strip() for part in re.split(r"[,;]", text) if part.strip()]


def add_category_text(categories: object, category: str) -> str:
    parts = category_parts(categories)
    if category and category.casefold() not in {part.casefold() for part in parts}:
        parts.append(category)
    return ", ".join(parts)
