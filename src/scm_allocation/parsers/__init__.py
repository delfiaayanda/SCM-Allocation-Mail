"""Allocation extraction parsers."""

from scm_allocation.parsers.excel import parse_excel
from scm_allocation.parsers.html import parse_html, parse_html_tables

__all__ = ["parse_excel", "parse_html", "parse_html_tables"]
