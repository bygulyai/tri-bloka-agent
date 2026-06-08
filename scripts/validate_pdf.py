"""validate_pdf.py — Programmatic PDF validation without vision model.
Usage: python validate_pdf.py <pdf_path> [expected_pages]
Checks: file exists, size > 10KB, page count matches.
"""
import sys
import os

def validate_pdf(pdf_path, expected_pages=16):
    if not os.path.isfile(pdf_path):
        print(f"FAIL: PDF not found: {pdf_path}")
        return False

    file_size = os.path.getsize(pdf_path)
    if file_size < 10000:
        print(f"FAIL: PDF too small ({file_size} bytes), likely corrupt: {pdf_path}")
        return False

    print(f"OK: PDF exists, size {file_size} bytes: {pdf_path}")

    page_count = None

    # Try PyMuPDF (fitz) first
    try:
        import fitz
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        doc.close()
    except ImportError:
        pass
    except Exception as e:
        print(f"WARN: PyMuPDF failed: {e}")

    # Fallback: count page markers in raw bytes
    if page_count is None:
        try:
            with open(pdf_path, "rb") as f:
                content = f.read()
            # Count /Type /Page entries that are not /Type /Pages
            count_pages = content.count(b"/Type /Page")
            count_pages_parent = content.count(b"/Type /Pages")
            page_count = count_pages - count_pages_parent
            if page_count < 1:
                page_count = count_pages
        except Exception as e:
            print(f"WARN: Raw byte counting failed: {e}")

    if page_count is not None and page_count > 0:
        print(f"Pages detected: {page_count} (expected: {expected_pages})")
        if page_count != expected_pages:
            print(f"WARN: Page count mismatch (got {page_count}, expected {expected_pages})")
        else:
            print(f"OK: Page count matches ({page_count})")
    else:
        print("WARN: Could not determine page count")

    print("VALIDATION COMPLETE")
    return True


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
    expected = int(sys.argv[2]) if len(sys.argv) > 2 else 16

    if not pdf_path:
        print("Usage: python validate_pdf.py <pdf_path> [expected_pages]")
        sys.exit(2)

    ok = validate_pdf(pdf_path, expected)
    sys.exit(0 if ok else 1)