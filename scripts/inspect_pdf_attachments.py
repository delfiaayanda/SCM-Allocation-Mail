"""Detailed inspector for the sample PDF attachments."""

from pathlib import Path
import pypdf


def inspect_pdfs(pdf_dir: Path):
    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    print("=" * 80)
    print(f"FOUND {len(pdf_files)} PDF FILES IN {pdf_dir}")
    print("=" * 80)

    for idx, pdf_path in enumerate(pdf_files, start=1):
        print(f"\n[{idx}] INSPECTING PDF: {pdf_path.name} ({pdf_path.stat().st_size} bytes)")
        print("-" * 60)
        try:
            reader = pypdf.PdfReader(pdf_path)
            num_pages = len(reader.pages)
            print(f"Number of pages: {num_pages}")
            
            # Check metadata
            meta = reader.metadata
            print(f"PDF Producer: {meta.producer if meta else 'None'}, Creator: {meta.creator if meta else 'None'}")

            total_text_length = 0
            all_text = []
            for page_idx, page in enumerate(reader.pages, start=1):
                text = page.extract_text() or ""
                total_text_length += len(text)
                all_text.append(text)

            is_text_based = total_text_length > 50
            print(f"Text-based: {'YES' if is_text_based else 'NO (likely scan/image)'}")
            print(f"Total extracted text length: {total_text_length} characters")

            # Print first page text preview
            if all_text:
                lines = [line.strip() for line in all_text[0].splitlines() if line.strip()]
                print(f"Extracted lines on page 1 ({len(lines)} lines):")
                for l_idx, line in enumerate(lines[:30], start=1):
                    print(f"  L{l_idx:02d}: {line}")
                if len(lines) > 30:
                    print(f"  ... [{len(lines) - 30} more lines on page 1]")
        except Exception as err:
            print(f"[ERROR] Failed reading PDF: {err}")


if __name__ == "__main__":
    p = Path("tests/fixtures/attachments")
    inspect_pdfs(p)
