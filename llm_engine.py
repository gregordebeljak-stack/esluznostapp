"""
Uporabi NVIDIA API (NVIDIA NIM - https://build.nvidia.com) - gostovan (cloud),
OpenAI-kompatibilen dostop do velikega števila modelov (Meta Llama, NVIDIA
Nemotron, DeepSeek, Mistral, Google Gemma ...), da iz surovega besedila PDF
izpiska iz zemljiške knjige (eZK) izlušči natančno pet podatkov o glavni
nepremičnini in njenem lastniku:

    ime_priimek | naslov | katastrska_obcina | parcelna_stevilka | delez

Ti podatki se nato v matching_engine.py uporabijo za zamenjavo ustreznih
rdečih polj v .docx predlogi.

Za delovanje potrebujete API ključ za NVIDIA API (prijavite se na
https://build.nvidia.com, odprite poljuben model in kliknite "Get API Key" -
ključ se začne z "nvapi-"), nastavljen kot okoljska spremenljivka
NVIDIA_API_KEY (npr. prek lokalne .env datoteke) - ali vnesen neposredno v
nadzorni plošči.

OPOMBA glede naslova (base URL): NVIDIA API privzeto streže OpenAI-
kompatibilen API na https://integrate.api.nvidia.com/v1. Če bi v prihodnje
uporabljali drugačen (npr. lastno gostovan NIM strežnik), ga vpišite v polje
"Naslov NVIDIA API (base URL)" v nadzorni plošči ali nastavite okoljsko
spremenljivko NVIDIA_BASE_URL - privzeta vrednost spodaj
(DEFAULT_NVIDIA_BASE_URL) je zato lahko potrebno prilagoditi.
"""

from __future__ import annotations

import json
import os
import re
import time
from typing import Dict, List

import httpx

# Privzeti naslov NVIDIA API-ja - glej opombo v glavi datoteke zgoraj;
# prilagodljivo prek okoljske spremenljivke NVIDIA_BASE_URL ali v UI-ju, če
# uporabljate npr. lastno gostovan NIM strežnik.
DEFAULT_NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"

# Privzeti naslov Google Gemini API-ja (OpenAI-kompatibilen endpoint) -
# prilagodljivo prek okoljske spremenljivke GEMINI_BASE_URL ali v UI-ju.
DEFAULT_GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai/"

# ----------------------------------------------------------------------------
# PODPORA VEČ PONUDNIKOM (NVIDIA / Gemini) - obe API-ja sta OpenAI-kompatibilni,
# zato ista extract_land_registry_data() spodaj deluje za oba - razlika je
# samo v naslovu (base_url), obliki API ključa in imenu, ki se prikaže v UI-ju.
# ----------------------------------------------------------------------------
PROVIDER_NVIDIA = "nvidia"
PROVIDER_GEMINI = "gemini"

PROVIDER_LABELS = {
    PROVIDER_NVIDIA: "NVIDIA API",
    PROVIDER_GEMINI: "Gemini API",
}

PROVIDER_ENV_KEY_VARS = {
    PROVIDER_NVIDIA: "NVIDIA_API_KEY",
    PROVIDER_GEMINI: "GEMINI_API_KEY",
}

PROVIDER_ENV_BASE_URL_VARS = {
    PROVIDER_NVIDIA: "NVIDIA_BASE_URL",
    PROVIDER_GEMINI: "GEMINI_BASE_URL",
}

PROVIDER_DEFAULT_BASE_URLS = {
    PROVIDER_NVIDIA: DEFAULT_NVIDIA_BASE_URL,
    PROVIDER_GEMINI: DEFAULT_GEMINI_BASE_URL,
}

PROVIDER_FALLBACK_MODELS = {
    PROVIDER_NVIDIA: [
        "meta/llama-3.3-70b-instruct",
        "deepseek-ai/deepseek-r1",
        "nvidia/llama-3.3-nemotron-super-49b-v1.5",
    ],
    PROVIDER_GEMINI: [
        "models/gemini-2.5-flash",
        "models/gemini-2.5-pro",
        "models/gemini-2.0-flash",
    ],
}


def detect_provider(api_key: str = "", base_url: str = "") -> str:
    """Samodejno prepozna ponudnika (NVIDIA ali Gemini) glede na naslov
    (base_url ima prednost - to je najbolj zanesljiv pokazatelj, kam gredo
    dejanske zahteve) in/ali obliko API ključa (Gemini ključi se začnejo z
    "AIza", NVIDIA ključi z "nvapi-"). Privzeto (če ni nobenega znaka) vrne
    NVIDIA."""
    base_url_l = (base_url or "").strip().lower()
    api_key_s = (api_key or "").strip()

    if "generativelanguage.googleapis.com" in base_url_l or "gemini" in base_url_l:
        return PROVIDER_GEMINI
    if "nvidia" in base_url_l:
        return PROVIDER_NVIDIA

    if api_key_s.lower().startswith("aiza"):
        return PROVIDER_GEMINI
    if api_key_s.startswith("nvapi-"):
        return PROVIDER_NVIDIA

    return PROVIDER_NVIDIA


def base_url_for_provider(provider: str) -> str:
    return PROVIDER_DEFAULT_BASE_URLS.get(provider, DEFAULT_NVIDIA_BASE_URL)


def _chat_completions_url(base_url: str) -> str:
    base_url = (base_url or DEFAULT_NVIDIA_BASE_URL).rstrip("/")
    return f"{base_url}/chat/completions"


# meta/llama-3.3-70b-instruct: zmogljiv, brezplačen model iz NVIDIA API
# kataloga, primeren za ekstrakcijo strukturiranih podatkov iz (lahko dolgih)
# pravnih PDF izpiskov. Za zahtevnejše/dvoumne primere (npr. solastništvo z
# več lastniki) lahko uporabnik v UI-ju izbere kateri koli drug model slug iz
# kataloga (glej seznam na https://build.nvidia.com/models), npr.
# "deepseek-ai/deepseek-r1" ali "nvidia/llama-3.3-nemotron-super-49b-v1.5".
DEFAULT_MODEL = "meta/llama-3.3-70b-instruct"

LAND_REGISTRY_KEYS = ["ime_priimek", "naslov", "katastrska_obcina", "parcelna_stevilka", "delez"]

__all__ = [
    "LLMExtractionError",
    "extract_land_registry_data",
    "count_probable_owners",
    "test_network_connectivity",
    "fetch_available_models",
    "LAND_REGISTRY_KEYS",
    "DEFAULT_MODEL",
    "DEFAULT_NVIDIA_BASE_URL",
    "DEFAULT_GEMINI_BASE_URL",
    "PROVIDER_NVIDIA",
    "PROVIDER_GEMINI",
    "PROVIDER_LABELS",
    "PROVIDER_ENV_KEY_VARS",
    "PROVIDER_ENV_BASE_URL_VARS",
    "PROVIDER_DEFAULT_BASE_URLS",
    "PROVIDER_FALLBACK_MODELS",
    "detect_provider",
    "base_url_for_provider",
]

# Vsak "imetnik" (lastnik) v eZK izpisku ima natanko eno vrstico "osebno ime:" -
# to štetje uporabimo kot NEODVISNO (od LLM-ja) preverbo, koliko lastnikov
# pričakujemo, da bi v UI-ju lahko opozorili, če jih je model izluščil manj
# (npr. pri solastnikih, razdeljenih v ločene razdelke "osnovni pravni
# položaj", glej opombo v SYSTEM_PROMPT spodaj).
_OSEBNO_IME_RE = re.compile(r"osebno\s+ime\s*:", re.IGNORECASE)


def count_probable_owners(pdf_text: str) -> int:
    """Prešteje verjetno število lastnikov v surovem besedilu PDF izpiska
    (po številu pojavitev "osebno ime:") - namenjeno kot varovalka v UI-ju,
    NEODVISNA od LLM ekstrakcije, da uporabnik opazi, če je model izluščil
    manj lastnikov, kot jih besedilo dejansko vsebuje."""
    return len(_OSEBNO_IME_RE.findall(pdf_text or ""))


SYSTEM_PROMPT = """Deluješ kot strokovni pravni asistent za nepremičninsko pravo v Sloveniji. Tvoja naloga je iz PDF izpiska iz zemljiške knjige (eZK) izluščiti natančno pet podatkov o GLAVNI NEPREMIČNINI in njenem LASTNIKU (osnovni pravni položaj / vknjižena lastninska pravica) - NE podatkov o imetnikih izvedenih pravic/služnosti (npr. Elektro Maribor, Občina, Vodovod ...), saj so to upravičenci bremen, ne lastniki nepremičnine.

Izlušči TOČNO te podatke:
- ime_priimek: ime in priimek lastnika (imetnika osnovnega pravnega položaja), natančno kot je zapisano (npr. "Miran Vajda")
- naslov: naslov lastnika, natančno kot je zapisan (npr. "Ruška cesta 146, 2345 Bistrica ob Dravi")
- katastrska_obcina: šifra in ime katastrske občine GLAVNE nepremičnine, natančno kot je zapisano (npr. "664 BISTRICA PRI RUŠAH")
- parcelna_stevilka: številka parcele GLAVNE nepremičnine, natančno kot je zapisana (npr. "177/4")
- delez: lastninski delež lastnika, natančno kot je zapisan (npr. "1/1", "1/2")

PRAVILA:
- Bodi izjemno natančen pri prepisovanju številk in imen - prepiši jih dobesedno.
- Če katerega od petih podatkov v izpisku zanesljivo ne najdeš, vrni zanj prazen niz "".
- Ne izmišljuj si podatkov in ne uganjaj, če nisi prepričan.
- Če je lastnikov več (solastništvo), vrni VSE lastnike v seznamu "owners".
- POMEMBNO - solastniki v LOČENIH razdelkih: pri parcelah v solastništvu eZK pogosto izpiše vsakega solastnika v SVOJEM ločenem razdelku "Osnovni pravni položaj nepremičnine" (vsak s svojim "ID osnovnega položaja", a z isto parcelno številko in ujemajočim se deležem, npr. en razdelek z deležem "1/2" in drug razdelek prav tako z deležem "1/2"). Ti razdelki so lahko med seboj ločeni z vmesnim razdelkom "omejitve" (npr. vknjižena služnost). To NISO ločene nepremičnine ali nepovezani pravni položaji, temveč SOLASTNIKI ISTE nepremičnine - obravnavaj jih kot en sklop in JIH VSE vključi v seznam "owners", tudi če so v besedilu ločeni z vmesnimi razdelki.
- Preden odgovoriš, preštej VSE pojavitve besedne zveze "osebno ime:" v besedilu - vsaka pojavitev pomeni enega lastnika (imetnika), ki ga MORAŠ vključiti kot ločen element v seznamu "owners".
- Vsak element v seznamu "owners" mora imeti TOČNO pet ključe: ime_priimek, naslov, katastrska_obcina, parcelna_stevilka, delez.
- Odgovori IZKLJUČNO z veljavnim JSON objektom, z top-level ključem "owners" in seznamom lastnikov. Brez uvodnega besedila, brez pojasnil."""


class LLMExtractionError(Exception):
    pass


def test_network_connectivity(proxy_url: str = None, base_url: str = None) -> str:
    """Preprost, hiter preizkus dosegljivosti NVIDIA API - neodvisen od
    veljavnosti API ključa. Uporabno za ločevanje omrežnih napak od napak ključa."""
    base_url = base_url or os.environ.get("NVIDIA_BASE_URL") or DEFAULT_NVIDIA_BASE_URL

    if proxy_url:
        try:
            with httpx.Client(proxy=proxy_url, timeout=8.0) as client:
                resp = client.get(base_url)
            return f"✅ Povezava preko proxyja uspešna (HTTP {resp.status_code})."
        except Exception as e:
            return f"❌ Povezava preko proxyja ({proxy_url}) ni uspela: {e}"

    try:
        start = time.time()
        with httpx.Client(timeout=8.0) as client:
            resp = client.get(base_url)
        elapsed_ms = int((time.time() - start) * 1000)
        return (
            f"✅ Povezava do NVIDIA API ({base_url}) uspešna (HTTP {resp.status_code}, "
            f"{elapsed_ms} ms)."
        )
    except (httpx.ConnectError, httpx.ConnectTimeout):
        return (
            f"❌ NVIDIA API ni dosegljiv na {base_url}. Preverite internetno povezavo / DNS "
            "nastavitve / VPN, in da naslov v polju 'Naslov NVIDIA API (base URL)' ustreza "
            "naslovu iz vašega NVIDIA API računa (glej https://docs.api.nvidia.com/nim)."
        )
    except httpx.TimeoutException:
        return f"❌ Povezava do NVIDIA API ({base_url}) se je iztekla (timeout)."
    except Exception as e:
        return f"❌ Povezava do NVIDIA API ({base_url}) ni uspela: {e}"


def fetch_available_models(
    api_key: str = None,
    base_url: str = None,
    proxy_url: str = None,
) -> List[str]:
    """Pridobi seznam razpoložljivih model slug-ov iz NVIDIA API (standarden
    OpenAI-kompatibilen GET {base_url}/models). Uporablja se za samodejno
    polnjenje spustnega seznama modelov v nadzorni plošči, takoj ko uporabnik
    vnese/spremeni naslov (in po možnosti API ključ).

    Vrne SORTIRAN seznam nizov (model ID-jev). Ob napaki (manjkajoč/napačen
    ključ, nedosegljiv strežnik, nepričakovana oblika odgovora ...) vrže
    LLMExtractionError - klicatelj naj to ujame in seznam po potrebi
    nadomesti s privzetim/ročnim vnosom, namesto da aplikacija zaradi tega
    odpove."""
    api_key = api_key or os.environ.get("NVIDIA_API_KEY")
    base_url = base_url or os.environ.get("NVIDIA_BASE_URL") or DEFAULT_NVIDIA_BASE_URL
    base_url = base_url.rstrip("/")
    url = f"{base_url}/models"

    headers = {}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    client_kwargs = {"timeout": 15.0}
    if proxy_url:
        client_kwargs["proxy"] = proxy_url

    try:
        with httpx.Client(**client_kwargs) as client:
            response = client.get(url, headers=headers)
    except (httpx.ConnectError, httpx.ConnectTimeout, httpx.ProxyError) as e:
        raise LLMExtractionError(
            f"Seznama modelov ni bilo mogoče pridobiti - NVIDIA API ni dosegljiv na {url}: {e}"
        ) from e
    except httpx.TimeoutException as e:
        raise LLMExtractionError(
            f"Zahteva za seznam modelov ({url}) je potekla (timeout)."
        ) from e

    if response.status_code in (401, 403):
        raise LLMExtractionError(
            "Neveljaven ali nepooblaščen API ključ - seznama modelov ni mogoče pridobiti. "
            "Vnesite veljaven NVIDIA API ključ (nvapi-...) zgoraj."
        )
    if response.status_code != 200:
        raise LLMExtractionError(
            f"Strežnik je pri pridobivanju seznama modelov ({url}) vrnil napako "
            f"(koda {response.status_code})."
        )

    try:
        data = response.json()
    except json.JSONDecodeError as e:
        raise LLMExtractionError(
            f"Odgovor s seznamom modelov ({url}) ni veljaven JSON."
        ) from e

    # Standardna OpenAI-kompatibilna oblika: {"data": [{"id": "..."}, ...]} - a
    # nekateri gateway-ji vrnejo kar gol seznam, zato podpremo tudi to obliko.
    items = data.get("data") if isinstance(data, dict) else data
    if not isinstance(items, list):
        raise LLMExtractionError(
            f"Nepričakovana oblika odgovora za seznam modelov ({url})."
        )

    model_ids: List[str] = []
    for item in items:
        if isinstance(item, dict) and item.get("id"):
            model_ids.append(str(item["id"]))
        elif isinstance(item, str) and item.strip():
            model_ids.append(item.strip())

    if not model_ids:
        raise LLMExtractionError(
            f"NVIDIA API je vrnil prazen seznam modelov ({url})."
        )

    return sorted(set(model_ids))


def _normalize_owner_record(owner: object) -> Dict[str, str] | None:
    """Normalizira en lastnikovo izluščeno JSON-polje v pričakovano indeksno strukturo."""
    if not isinstance(owner, dict):
        return None
    return {key: str(owner.get(key) or "").strip() for key in LAND_REGISTRY_KEYS}


def _coerce_owner_records(payload: object) -> List[Dict[str, str]] | None:
    """Vzame t.i. JSON-odgovor modela in vrne seznam lastnikov, če je odgovor v enem od
    podprtih list-odgovornih oblik. V nasprotnem primeru vrne None.
    """
    if isinstance(payload, list):
        normalized = []
        for owner in payload:
            record = _normalize_owner_record(owner)
            if record is not None:
                normalized.append(record)
        return normalized or None

    if not isinstance(payload, dict):
        return None

    # Top-level owner-keys are the new contract, but we support variants as a compatibility layer.
    for owners_key in ("owners", "lastniki", "solastniki"):
        candidate = payload.get(owners_key)
        if isinstance(candidate, list):
            normalized = []
            for owner in candidate:
                record = _normalize_owner_record(owner)
                if record is not None:
                    normalized.append(record)
            return normalized or [{key: "" for key in LAND_REGISTRY_KEYS}]

    # Legacy: one owner object without a surrounding list.
    record = _normalize_owner_record(payload)
    if record and any(v for v in record.values()):
        return [record]
    return None


def extract_land_registry_data(
    pdf_text: str,
    api_key: str = None,
    proxy_url: str = None,
    model: str = DEFAULT_MODEL,
    base_url: str = None,
) -> Dict[str, str] | List[Dict[str, str]]:
    """
    pdf_text: surovo besedilo izpiska iz zemljiške knjige
    api_key: API ključ, ustvarjen na https://build.nvidia.com (Bearer token, oblike "nvapi-...")
    proxy_url: neobvezen HTTP(S) proxy, npr. "http://uporabnik:geslo@proxy.podjetje.si:8080"
    model: NVIDIA API "model slug", npr. "meta/llama-3.3-70b-instruct" ali "deepseek-ai/deepseek-r1"
    base_url: naslov NVIDIA API-ja (privzeto https://integrate.api.nvidia.com/v1 oz.
        NVIDIA_BASE_URL, če je nastavljena - glej opombo v glavi datoteke)
    Vrne: {"ime_priimek": ..., "naslov": ..., "katastrska_obcina": ..., "parcelna_stevilka": ..., "delez": ...}
    """
    api_key = api_key or os.environ.get("NVIDIA_API_KEY")
    base_url = base_url or os.environ.get("NVIDIA_BASE_URL") or DEFAULT_NVIDIA_BASE_URL
    if not api_key:
        raise LLMExtractionError(
            "Manjka NVIDIA_API_KEY. Nastavite ga kot okoljsko spremenljivko (najlažje prek "
            "lokalne .env datoteke) ali ga vnesite v nadzorni plošči (levi stolpec). "
            "Ključ ustvarite na https://build.nvidia.com (odprite poljuben model in kliknite "
            "'Get API Key')."
        )

    user_message = (
        "SUROVO BESEDILO IZ PDF IZPISKA ZEMLJIŠKE KNJIGE:\n\n" + pdf_text +
        "\n\nVrni JSON z top-level ključem 'owners' in seznamom lastnikov. "
        "Vsak lastnik naj vsebuje podatke: ime_priimek, naslov, katastrska_obcina, "
        "parcelna_stevilka, delez."
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ],
        "response_format": {"type": "json_object"},
        "max_tokens": 1000,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        # Neobvezna, a priporočena identifikacijska glava.
        "X-Title": "Nadzorna plosca za avtomatizacijo dokumentov",
    }

    client_kwargs = {"timeout": 60.0}
    if proxy_url:
        client_kwargs["proxy"] = proxy_url

    url = _chat_completions_url(base_url)

    try:
        with httpx.Client(**client_kwargs) as client:
            response = client.post(url, json=payload, headers=headers)
    except (httpx.ConnectError, httpx.ConnectTimeout, httpx.ProxyError) as e:
        raise LLMExtractionError(
            f"Ni bilo mogoče vzpostaviti povezave z NVIDIA API ({base_url}). "
            "To NI napaka API ključa, temveč omrežna težava na tem računalniku. Preverite:\n"
            "  1) internetno povezavo\n"
            "  2) da naslov v polju 'Naslov NVIDIA API (base URL)' ustreza naslovu iz vašega "
            f"NVIDIA API računa (trenutno: {base_url})\n"
            "  3) da požarni zid / protivirusni program ne blokira Python/Streamlit procesa\n"
            "  4) če ste za korporativnim proxyjem: izpolnite polje 'HTTP(S) proxy' v levem stolpcu\n"
            "  5) VPN - poskusite začasno izklopiti VPN\n"
            f"Izvirna napaka: {e}"
        ) from e
    except httpx.TimeoutException as e:
        raise LLMExtractionError(
            "Zahteva je potekla (timeout) - strežnik ni odgovoril pravočasno. "
            "Poskusite znova ali preverite hitrost/stabilnost povezave."
        ) from e

    if response.status_code in (401, 403):
        raise LLMExtractionError(
            "Neveljaven ali nepooblaščen NVIDIA_API_KEY (napaka avtentikacije). "
            "Preverite ključ na https://build.nvidia.com."
        )
    if response.status_code == 429:
        raise LLMExtractionError(
            "Presežena kvota/omejitev zahtev (429) pri NVIDIA API. Počakajte trenutek in "
            "poskusite znova, ali preverite stanje kredita/kvot na svojem NVIDIA API računu."
        )
    if response.status_code >= 500:
        raise LLMExtractionError(
            f"NVIDIA API strežnik je vrnil napako (koda {response.status_code}). "
            "Poskusite znova čez nekaj trenutkov."
        )
    if response.status_code != 200:
        try:
            err_body = response.json()
            err_msg = err_body.get("error", {}).get("message", response.text[:300])
        except Exception:
            err_msg = response.text[:300]
        raise LLMExtractionError(f"NVIDIA API je vrnil napako (koda {response.status_code}): {err_msg}")

    try:
        data = response.json()
    except json.JSONDecodeError as e:
        raise LLMExtractionError(
            f"NVIDIA API ni vrnil veljavnega JSON odgovora:\n{response.text[:500]}"
        ) from e

    if "error" in data:
        err = data["error"]
        msg = err.get("message", err) if isinstance(err, dict) else err
        raise LLMExtractionError(f"NVIDIA API je vrnil napako: {msg}")

    try:
        raw_text = (data["choices"][0]["message"]["content"] or "").strip()
    except (KeyError, IndexError) as e:
        raise LLMExtractionError(
            f"Nepričakovana oblika odgovora NVIDIA API:\n{json.dumps(data)[:500]}"
        ) from e

    # Varnostna mreža, če model kljub JSON-mode konfiguraciji doda ``` ograjo.
    if raw_text.startswith("```"):
        raw_text = raw_text.strip("`")
        if raw_text.lower().startswith("json"):
            raw_text = raw_text[4:].strip()

    try:
        result = json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise LLMExtractionError(
            f"Model ni vrnil veljavnega JSON-a. Prejeti odgovor:\n{raw_text[:500]}"
        ) from e

    if not isinstance(result, dict):
        raise LLMExtractionError("Model ni vrnil JSON objekta (dict).")

    # Podprti odgovori so ali en sam lastnik (stari enojni format) ali
    # več lastnikov v top-level ključu "owners" / "lastniki" / "solastniki".
    owners = None
    for owners_key in ("owners", "lastniki", "solastniki"):
        candidate = result.get(owners_key)
        if isinstance(candidate, list):
            owners = candidate
            break

    if owners is not None:
        normalized = []
        for owner in owners:
            if not isinstance(owner, dict):
                continue
            normalized.append({key: str(owner.get(key) or "").strip() for key in LAND_REGISTRY_KEYS})
        if not normalized:
            normalized.append({key: "" for key in LAND_REGISTRY_KEYS})
        return normalized

    # Če model vrne korespondenco "owners" z več lastniki, jo normaliziramo
    # v seznam zapisov. Pri oblikah "legacy" z enim lastnikom vrnemo dict.
    owner_records = _coerce_owner_records(result)
    if owner_records:
        return owner_records

    # Poskrbi, da so navzoči vsi pričakovani ključi (manjkajoče nadomesti s "").
    return {key: str(result.get(key) or "").strip() for key in LAND_REGISTRY_KEYS}
