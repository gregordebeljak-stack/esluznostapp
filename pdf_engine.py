"""
pdf_engine.py
-------------
Ekstrakcija surovega besedila iz PDF izpiskov iz zemljiške knjige (eZK) s
pomočjo pdfplumber. Namerno vrača berljivo, a čim bolj "surovo" besedilo -
strukturiranje (kdo je lastnik, katera parcela, katera bremena...) prepustimo
LLM-ju v llm_engine.py, ki bolje razume prosto besedilo pravnih izpiskov kot
regex/pravila.
"""

from __future__ import annotations

import io
from typing import List

import pdfplumber


def extract_pdf_text(file_bytes: bytes) -> str:
    """Izvleče besedilo iz vseh strani PDF-ja, ohranja vrstni red strani."""
    pages_text: List[str] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages_text.append(f"--- Stran {i + 1} ---\n{text}")
    return "\n\n".join(pages_text)


def extract_pdf_tables(file_bytes: bytes) -> List[List[List[str]]]:
    """Pomožna funkcija: izvleče morebitne tabele (za napredno uporabo/debug)."""
    all_tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                all_tables.append(table)
    return all_tables
