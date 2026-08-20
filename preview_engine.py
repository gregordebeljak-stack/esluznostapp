"""
preview_engine.py
------------------
Ustvari pravo VIZUALNO primerjavo dokumenta pred in po izpolnitvi (original
z rdečimi polji vs. končni dokument s črnimi polji), stran ob strani.

Uporablja LibreOffice (soffice --headless) za pretvorbo .docx -> .pdf in
poppler-utils (pdftoppm) za pretvorbo .pdf -> .jpg. To je edini zanesljiv,
brezplačen način za popolnoma zvest prikaz Wordovega dokumenta (barve,
tabele, pisave) izven samega Worda.

Če LibreOffice/poppler nista nameščena, `is_available()` vrne False in
aplikacija namesto tega prikaže enostavnejši HTML predogled (glej app.py).
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

from PIL import Image, ImageDraw


def is_available() -> bool:
    return shutil.which("soffice") is not None and shutil.which("pdftoppm") is not None


def _docx_to_images(docx_bytes: bytes, workdir: Path, prefix: str) -> List[Path]:
    docx_path = workdir / f"{prefix}.docx"
    docx_path.write_bytes(docx_bytes)

    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(workdir), str(docx_path)],
        capture_output=True, text=True, timeout=90,
    )
    pdf_path = workdir / f"{prefix}.pdf"
    if not pdf_path.exists():
        raise RuntimeError(f"Pretvorba v PDF ni uspela: {result.stderr or result.stdout}")

    img_prefix = workdir / f"{prefix}_page"
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", "120", str(pdf_path), str(img_prefix)],
        capture_output=True, text=True, timeout=60, check=True,
    )
    images = sorted(workdir.glob(f"{prefix}_page-*.jpg"))
    if not images:
        # pdftoppm z eno stranjo včasih ne doda "-1" pripone
        images = sorted(workdir.glob(f"{prefix}_page*.jpg"))
    return images


def docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    """Pretvori .docx v .pdf (prek LibreOffice) in vrne surove PDF bajte -
    uporabno za neposreden predogled/tiskanje dokumenta v brskalniku, brez
    izvoza vmesnih slik po straneh."""
    if not is_available():
        raise RuntimeError(
            "LibreOffice ali poppler-utils (pdftoppm) nista nameščena na tem sistemu. "
            "Namestite ju za predogled/tiskanje PDF (glej README)."
        )
    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        docx_path = workdir / "print_doc.docx"
        docx_path.write_bytes(docx_bytes)
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(workdir), str(docx_path)],
            capture_output=True, text=True, timeout=90,
        )
        pdf_path = workdir / "print_doc.pdf"
        if not pdf_path.exists():
            raise RuntimeError(f"Pretvorba v PDF ni uspela: {result.stderr or result.stdout}")
        return pdf_path.read_bytes()


def build_side_by_side_comparison(
    original_docx_bytes: bytes,
    final_docx_bytes: bytes,
    max_pages: int = 6,
) -> List[bytes]:
    """Vrne seznam JPEG slik (kot bytes), po eno na stran dokumenta, kjer je
    original prikazan levo in izpolnjena verzija desno."""
    if not is_available():
        raise RuntimeError(
            "LibreOffice ali poppler-utils (pdftoppm) nista nameščena na tem sistemu. "
            "Namestite ju za vizualni predogled pred/po (glej README)."
        )

    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        orig_images = _docx_to_images(original_docx_bytes, workdir, "original")
        final_images = _docx_to_images(final_docx_bytes, workdir, "final")

        n_pages = min(len(orig_images), len(final_images), max_pages)
        combined_bytes: List[bytes] = []

        for i in range(n_pages):
            orig_img = Image.open(orig_images[i])
            final_img = Image.open(final_images[i])
            w, h = orig_img.size
            gap = 16
            label_h = 34
            combined = Image.new("RGB", (w * 2 + gap, h + label_h), "white")
            draw = ImageDraw.Draw(combined)
            draw.text((w // 2 - 45, 8), "PRED (original)", fill="black")
            draw.text((w + gap + w // 2 - 55, 8), "PO (izpolnjeno)", fill="black")
            combined.paste(orig_img, (0, label_h))
            combined.paste(final_img, (w + gap, label_h))

            import io
            buf = io.BytesIO()
            combined.save(buf, format="JPEG", quality=85)
            combined_bytes.append(buf.getvalue())

        return combined_bytes