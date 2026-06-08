import sys
from pathlib import Path

import fitz


def main() -> int:
    if len(sys.argv) < 4:
        print("Usage: render_pdf_pages.py <pdf> <outdir> <page1> [<page2> ...]", file=sys.stderr)
        return 2

    pdf_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    pages = [int(x) for x in sys.argv[3:]]

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    for page_num in pages:
        page = doc[page_num - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        out_path = out_dir / f"page_{page_num:02d}.png"
        pix.save(out_path)
        print(out_path)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
