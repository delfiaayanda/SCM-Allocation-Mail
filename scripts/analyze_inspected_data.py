"""Analyze inspected emails to extract structure, fields, and patterns."""

import json
from pathlib import Path
import re
from html.parser import HTMLParser


class TableExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tables = []
        self.current_table = []
        self.current_row = []
        self.current_cell = []
        self.in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.current_table = []
        elif tag == "tr":
            self.current_row = []
        elif tag in ("td", "th"):
            self.in_cell = True
            self.current_cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self.in_cell = False
            cell_text = "".join(self.current_cell).strip()
            cell_text = re.sub(r"\s+", " ", cell_text)
            self.current_row.append(cell_text)
        elif tag == "tr":
            if self.current_row:
                self.current_table.append(self.current_row)
                self.current_row = []
        elif tag == "table":
            if self.current_table:
                self.tables.append(self.current_table)
                self.current_table = []

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell.append(data)


def main():
    for i in range(1, 6):
        p = Path(f"data/inspection/email_{i:03d}.json")
        data = json.loads(p.read_text(encoding="utf-8"))
        print("=" * 70)
        print(f"EMAIL {i:03d}: {data['subject']}")
        print(f"From: {data['sender']} | Date: {data['received_time']}")
        print(f"Format: {data['body_format']}")
        print(f"Attachments: {[a['filename'] for a in data['attachments']]}")
        
        # Analyze plain text
        lines = [line.strip() for line in data["body_plain_text"].splitlines() if line.strip()]
        print(f"\n--- Plain text lines (first 10 non-empty) ---")
        for line in lines[:10]:
            print("  ", line)

        # Analyze HTML tables if any
        if data["has_html_table"]:
            parser = TableExtractor()
            parser.feed(data["body_html"])
            print(f"\n--- Detected {len(parser.tables)} HTML Table(s) ---")
            for t_idx, tbl in enumerate(parser.tables, start=1):
                print(f"  Table {t_idx} ({len(tbl)} rows):")
                for r_idx, row in enumerate(tbl[:5], start=1):
                    print(f"    Row {r_idx}: {row}")
                if len(tbl) > 5:
                    print(f"    ... [{len(tbl) - 5} more rows]")
        print("-" * 70)


if __name__ == "__main__":
    main()
