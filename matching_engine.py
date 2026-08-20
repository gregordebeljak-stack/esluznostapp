"""
matching_engine.py
-------------------
Prepozna, kateri od zaznanih rdečih polj v .docx predlogi ustrezajo petim
kategorijam podatkov iz zemljiške knjige:

    ime_priimek | naslov | katastrska_obcina | parcelna_stevilka | delez

in jih zamenja z vrednostmi, izluščenimi iz PDF izpiska. Polja, ki ne
ustrezajo nobeni od teh kategorij (npr. opis del, datumi, podatki druge
pogodbene stranke), ostanejo nedotaknjena - uporabnik jih po potrebi uredi
ročno na nadzorni plošči.

Prepoznavanje deluje v dveh korakih:
1. KONTEKST - iskanje ključnih besed v oznaki polja, besedilu pred/za njim
   in glavi tabelskega stolpca (npr. "parc. št.", "k.o.", "delež").
2. OBLIKA VREDNOSTI - regex prepoznavanje vzorcev (ime+priimek, naslov s
   poštno številko, ulomek za parcelo/delež, "šifra + VELIKE ČRKE" za k.o.).
   To deluje tudi, kadar je več podatkov združenih v enem rdečem odseku
   (npr. "Janez Novak, Slovenska cesta 1, 1000 Ljubljana, EMŠO: ").
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional

from docx_engine import Field, apply_field_value

LAND_REGISTRY_KEYS = ["ime_priimek", "naslov", "katastrska_obcina", "parcelna_stevilka", "delez"]

_NAME_RE = re.compile(
    r"^(?P<name>[A-ZČŠŽ][\wčšžĐđ.\-]*(?:\s+[A-ZČŠŽ][\wčšžĐđ.\-]*){1,3})\s*,\s*(?P<rest>.+)$"
)
_ADDRESS_RE = re.compile(
    r"(?P<addr>[^,]*\d+[a-zA-Z]?\s*,\s*\d{4}\s+[A-ZČŠŽ][\wčšžĐđ\- ]*)"
)
_KO_RE = re.compile(r"^\s*\d{2,4}\s+[A-ZČŠŽ][A-ZČŠŽ \.]+\s*$")
_FRACTION_RE = re.compile(r"(\d+\s*/\s*\d+)")
_BARE_NAME_RE = re.compile(
    r"^[A-ZČŠŽ][\wčšžĐđ.\-]*(\s+[A-ZČŠŽ][\wčšžĐđ.\-]*){1,3}$"
)


@dataclass
class MatchResult:
    field_id: str
    role: str
    old_value: str
    new_value: str
    changed: bool


def _context_roles(context_hint: str) -> set:
    text = context_hint.lower()
    roles = set()
    if "parc" in text:
        roles.add("parcelna_stevilka")
    if "delež" in text or "delez" in text:
        roles.add("delez")
    if "k.o." in text or "katastrsk" in text:
        roles.add("katastrska_obcina")
    if "naslov" in text:
        roles.add("naslov")
    if any(k in text for k in ["ime", "priimek", "lastnik", "imetnik", "zavezanec", "upravičenec", "podpis"]):
        roles.add("ime_priimek")
    return roles


def _preserve_surname_case(old_value: str, new_value: str) -> str:
    """Če je bil v originalu priimek zapisan z VELIKIMI črkami (npr. slogovna
    konvencija podpisnega polja), enako obliko ohrani tudi v novi vrednosti."""
    old_tokens = old_value.strip().split()
    if len(old_tokens) >= 2 and old_tokens[-1].isupper() and len(old_tokens[-1]) > 1:
        new_tokens = new_value.strip().split()
        if new_tokens:
            new_tokens[-1] = new_tokens[-1].upper()
            return " ".join(new_tokens)
    return new_value


def _build_replacement(field: Field, extracted: Dict[str, str]) -> Optional[tuple]:
    """Vrne (vloga, nova_vrednost) ali None, če polje ne ustreza nobeni od
    petih kategorij."""
    old_value = field.value
    roles = _context_roles(field.context_hint)

    # 1) Kombinacija: ime+priimek na začetku, ločeno z vejico, sledi naslov
    #    (in morda še dodatno besedilo, npr. ", EMŠO: ", ki ga ohranimo).
    m = _NAME_RE.match(old_value)
    if m and extracted.get("ime_priimek") and extracted.get("naslov"):
        rest = m.group("rest")
        addr_m = _ADDRESS_RE.search(rest)
        if addr_m:
            trailing = rest[addr_m.end():]  # npr. ", EMŠO: "
            new_name = _preserve_surname_case(m.group("name"), extracted["ime_priimek"])
            new_val = f"{new_name}, {extracted['naslov']}{trailing}"
            return ("ime_priimek+naslov", new_val)

    # 2) Ulomek (parcela ali delež) - kontekst (npr. glava stolpca "delež") odloči,
    #    kateri od obeh gre za kaj. To preverimo PRED katastrsko občino, ker lahko
    #    besedilo "k.o." iz sosednjega besedila (v istem odstavku) sicer napačno
    #    "prepusti" vlogo katastrska_obcina polju za parcelno številko.
    frac_m = _FRACTION_RE.search(old_value)
    if frac_m:
        if "delez" in roles and extracted.get("delez"):
            new_val = old_value[:frac_m.start()] + extracted["delez"] + old_value[frac_m.end():]
            return ("delez", new_val)
        if extracted.get("parcelna_stevilka"):
            new_val = old_value[:frac_m.start()] + extracted["parcelna_stevilka"] + old_value[frac_m.end():]
            return ("parcelna_stevilka", new_val)

    # 3) Katastrska občina: prepoznamo IZKLJUČNO po obliki "664 BISTRICA PRI RUŠAH"
    #    (šifra + velike črke) - namerno NE po ključni besedi "k.o." v kontekstu,
    #    ker se ta pogosto pojavi tudi tik ob sosednjem polju za parcelno številko.
    if _KO_RE.match(old_value) and extracted.get("katastrska_obcina"):
        return ("katastrska_obcina", extracted["katastrska_obcina"])

    # 4) Samo naslov (brez imena spredaj), prepoznan po obliki "ulica št., poštna_st kraj"
    if _ADDRESS_RE.search(old_value) and not _KO_RE.match(old_value):
        if extracted.get("naslov") and ("naslov" in roles or _ADDRESS_RE.fullmatch(old_value.strip())):
            return ("naslov", extracted["naslov"])

    # 5) Samo ime in priimek (brez naslova zraven), prepoznano po obliki ali kontekstu
    if extracted.get("ime_priimek") and (_BARE_NAME_RE.match(old_value.strip()) or "ime_priimek" in roles):
        # Izogni se lažnim ujemanjem s katastrsko občino (VELIKE ČRKE + šifra) - že obravnavano zgoraj
        if not _KO_RE.match(old_value):
            new_val = _preserve_surname_case(old_value, extracted["ime_priimek"])
            return ("ime_priimek", new_val)

    return None


def classify_role(field: Field) -> Optional[str]:
    """Bralno-only različica prepoznavanja vloge polja (brez zamenjave vrednosti) -
    uporabna za prikaz v UI-ju, tudi če je bil dokument med tem ponovno razčlenjen
    (npr. po dodajanju vrstic za več parcel, ko se id-ji polj spremenijo)."""
    old_value = field.value
    roles = _context_roles(field.context_hint)

    if _NAME_RE.match(old_value) and _ADDRESS_RE.search(_NAME_RE.match(old_value).group("rest")):
        return "ime_priimek+naslov"
    frac_m = _FRACTION_RE.search(old_value)
    if frac_m:
        return "delez" if "delez" in roles else "parcelna_stevilka"
    if _KO_RE.match(old_value):
        return "katastrska_obcina"
    if _ADDRESS_RE.search(old_value) and not _KO_RE.match(old_value):
        if "naslov" in roles or _ADDRESS_RE.fullmatch(old_value.strip()):
            return "naslov"
    if _BARE_NAME_RE.match(old_value.strip()) or "ime_priimek" in roles:
        if not _KO_RE.match(old_value):
            return "ime_priimek"
    return None


def match_and_apply(fields: Dict[str, Field], extracted: Dict[str, str]) -> List[MatchResult]:
    """Za vsako polje ugotovi, ali ustreza eni od 5 kategorij podatkov iz
    zemljiške knjige; če da, zamenja vrednost (in obarva črno prek
    apply_field_value) ter vrne poročilo o vseh ujemanjih (za prikaz
    uporabniku, tudi če se vrednost ni spremenila - polje je bilo 'potrjeno')."""
    results: List[MatchResult] = []
    for fid, f in fields.items():
        if f.is_blank_placeholder:
            continue
        match = _build_replacement(f, extracted)
        if match is None:
            continue
        role, new_value = match
        old_value = f.value
        changed = new_value.strip() != old_value.strip()
        apply_field_value(f, new_value)
        results.append(MatchResult(field_id=fid, role=role, old_value=old_value, new_value=new_value, changed=changed))
    return results
