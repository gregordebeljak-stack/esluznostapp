"""
owner_grouping_engine.py
-------------------------
Združevanje izluščenih podatkov iz VEČ (do 30) PDF izpiskov zemljiške knjige
po lastnikih - omogoča pregled "kateremu lastniku pripadajo katere parcele
in kakšni so njegovi/njeni deleži", brez izpolnjevanja .docx predloge.

Ne dotika se .docx logike - deluje izključno na že izluščenih podatkih
(seznamu slovarjev s ključi ime_priimek/naslov/katastrska_obcina/
parcelna_stevilka/delez + file_name), ki jih za vsak PDF vrne
llm_engine.extract_land_registry_data.
"""

from __future__ import annotations

import re
import unicodedata
from collections import OrderedDict
from typing import Dict, List


def _normalize_key(text: str) -> str:
    """Normalizira besedilo za PRIMERJAVO (ne za prikaz) - odstrani šumnike,
    podvojene presledke in velike/male črke, da se npr. 'Janez Novak' in
    'JANEZ  NOVAK' prepoznata kot ista oseba."""
    text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def group_by_owner(records: List[dict]) -> "OrderedDict[str, dict]":
    """records: seznam slovarjev {"file_name":, "ime_priimek":, "naslov":,
    "katastrska_obcina":, "parcelna_stevilka":, "delez":}, po eden na PDF.

    Vrne OrderedDict {owner_key: {"ime_priimek":, "naslov":, "parcels": [...]}}.
    Lastnika prepoznamo po paru (ime_priimek, naslov) - če se isto ime pojavi
    z istim naslovom v več PDF-jih, se vse parcele združijo pod enega lastnika."""
    owners: "OrderedDict[str, dict]" = OrderedDict()
    for rec in records:
        name = (rec.get("ime_priimek") or "").strip()
        addr = (rec.get("naslov") or "").strip()
        key = _normalize_key(name) + "||" + _normalize_key(addr)
        if key not in owners:
            owners[key] = {
                "ime_priimek": name or "(ni prepoznano)",
                "naslov": addr,
                "parcels": [],
            }
        owners[key]["parcels"].append({
            "katastrska_obcina": rec.get("katastrska_obcina", ""),
            "parcelna_stevilka": rec.get("parcelna_stevilka", ""),
            "delez": rec.get("delez", ""),
            "source_file": rec.get("file_name", ""),
        })
    return owners


def group_by_ownership_unit(records: List[dict]) -> "OrderedDict[str, dict]":
    """Podobno kot `group_by_owner`, a združuje SOLASTNIKE (npr. zakonca, ki
    si delita isto parcelo, vsak s svojim deležem) v ENO SKUPNO "lastniško
    enoto", namesto da bi vsak nastopal kot svoj ločen "lastnik" - tako se
    npr. Branko in Marija Cafuta, ki oba nastopata v istem PDF izpisku za
    parcelo 951/8 (vsak z deležem 1/2), prikažeta kot ENA enota "Branko
    Cafuta, Marija Cafuta", ne kot dva ločena vnosa.

    Osebi štejeta za solastnika, če se pojavita SKUPAJ v istem izvornem PDF-ju
    (isti `file_name`) - to je zanesljiv znak, da gre za solastnika iste
    parcele/izpiska. Če se ista oseba pojavi v VEČ datotekah (npr. lastnik
    A je solastnik parcele 1 skupaj z B, in solastnik parcele 2 skupaj s C),
    se prek te osebe med sabo združita tudi enoti B in C (union-find) - tako
    dobimo skupine vseh med sabo solastniško povezanih oseb.

    Vrne OrderedDict {unit_key: {"owners": [{"ime_priimek":, "naslov":}, ...]
    (unikatne osebe enote, v vrstnem redu prvega pojava),
    "records": [...vsi izvirni zapisi (oseba+parcela) te enote, v vrstnem
    redu prvega pojava...]}}."""

    def person_key(rec: dict) -> str:
        name = (rec.get("ime_priimek") or "").strip()
        addr = (rec.get("naslov") or "").strip()
        return _normalize_key(name) + "||" + _normalize_key(addr)

    parent: Dict[str, str] = {}

    def find(x: str) -> str:
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: str, b: str):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    # 1) Registriraj vse osebe, nato poveži tiste, ki nastopajo v ISTI
    #    izvorni datoteki (solastniki iste parcele/izpiska).
    by_file: Dict[str, List[str]] = {}
    for rec in records:
        pk = person_key(rec)
        find(pk)  # registracija v union-find
        fname = rec.get("file_name", "")
        by_file.setdefault(fname, []).append(pk)

    for keys in by_file.values():
        first = keys[0]
        for k in keys[1:]:
            union(first, k)

    # 2) Razporedi zapise po korenski (union-find) skupini, ohrani vrstni red
    #    prvega pojava skupine in prvega pojava vsake osebe znotraj nje.
    units: "OrderedDict[str, dict]" = OrderedDict()
    seen_owners_per_unit: Dict[str, set] = {}
    for rec in records:
        pk = person_key(rec)
        root = find(pk)
        if root not in units:
            units[root] = {"owners": [], "records": []}
            seen_owners_per_unit[root] = set()
        if pk not in seen_owners_per_unit[root]:
            seen_owners_per_unit[root].add(pk)
            units[root]["owners"].append({
                "ime_priimek": (rec.get("ime_priimek") or "").strip() or "(ni prepoznano)",
                "naslov": (rec.get("naslov") or "").strip(),
            })
        units[root]["records"].append(rec)

    return units


def group_by_katastrska_obcina(records: List[dict]) -> "OrderedDict[str, dict]":
    """Enako kot group_by_owner, a združeno po katastrski občini - uporabno za
    vprašanje 'kateri lastniki imajo parcele v tej k.o. in s kakšnim deležem?'"""
    groups: "OrderedDict[str, dict]" = OrderedDict()
    for rec in records:
        ko = (rec.get("katastrska_obcina") or "").strip() or "(ni prepoznano)"
        if ko not in groups:
            groups[ko] = {"katastrska_obcina": ko, "entries": []}
        groups[ko]["entries"].append({
            "ime_priimek": rec.get("ime_priimek", ""),
            "naslov": rec.get("naslov", ""),
            "parcelna_stevilka": rec.get("parcelna_stevilka", ""),
            "delez": rec.get("delez", ""),
            "source_file": rec.get("file_name", ""),
        })
    return groups


def find_name_conflicts(records: List[dict]) -> List[str]:
    """Vrne informativna opozorila, če se isto ime pojavi z RAZLIČNIMI naslovi
    (možna napaka pri branju/izbiri datotek, ali dve različni osebi z istim
    imenom) - samo v obvestilo uporabniku, ne vpliva na razvrščanje."""
    warnings: List[str] = []
    by_name: Dict[str, set] = {}
    display_name: Dict[str, str] = {}
    for rec in records:
        name = (rec.get("ime_priimek") or "").strip()
        addr = (rec.get("naslov") or "").strip()
        if not name:
            continue
        norm = _normalize_key(name)
        display_name.setdefault(norm, name)
        by_name.setdefault(norm, set()).add(addr)
    for norm, addrs in by_name.items():
        addrs_nonempty = {a for a in addrs if a}
        if len(addrs_nonempty) > 1:
            warnings.append(
                f"Ime '{display_name[norm]}' se pojavi z več različnimi naslovi: "
                f"{', '.join(sorted(addrs_nonempty))} - preverite, ali gre za isto osebo."
            )
    return warnings


def to_csv_rows(records: List[dict]) -> List[List[str]]:
    """Pripravi ploščato tabelo (za CSV izvoz) - ena vrstica na parcelo/PDF."""
    header = ["Ime in priimek", "Naslov", "Katastrska občina", "Parcelna številka", "Delež", "Izvorna datoteka"]
    rows = [header]
    for rec in records:
        rows.append([
            rec.get("ime_priimek", ""),
            rec.get("naslov", ""),
            rec.get("katastrska_obcina", ""),
            rec.get("parcelna_stevilka", ""),
            rec.get("delez", ""),
            rec.get("file_name", ""),
        ])
    return rows
