"""
Avtomatizacija prijave v portal eSodstvo (evlozisce/esodisce.si) s SI-PASS
kvalificiranim digitalnim potrdilom. Ker je potrdilo že naloženo na računalnik,
skripta odpre vidno okno brskalnika za prijavo, nato pa stanje (piškotke)
shrani za nadaljnje iskanje izpiskov.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

SI_PASS_LOGIN_URL = (
    "https://esodisce.si/vash-keycloak/realms/vash/protocol/openid-connect/auth"
    "?client_id=vash&response_type=code&scope=openid"
    "&redirect_uri=https://esodisce.si/spa/auth/callback"
    "&kc_idp_hint=si-pass-prod&audience=vash&kc_locale=sl&ui_locales=sl"
)
ACCESS_URL = "https://esodisce.si/vash/ivs/access/ezk"
SEARCH_URL = "https://esodisce.si/evlozisce/javni_izpisi/list.html#"
ESODISCE_ORIGIN = "https://esodisce.si"

LOGIN_TIMEOUT_MS = 120_000       # Prijava lahko vzame čas (uporabnik mora klikniti certifikat)
NAV_TIMEOUT_MS = 30_000
DOWNLOAD_TIMEOUT_MS = 30_000

# Blokirani viri po tipu - CSS/slike/pisave/mediji za samo iskanje in prenos
# PDF-ja niso potrebni (stran se izriše brez njih, jQuery UI widget deluje
# povsem enako). Dodatno blokiramo še znane domene za analitiko/sledenje
# (Google Analytics/Tag Manager, Hotjar, Facebook Pixel ipd.) - te dodatne
# klice esodisce.si sicer sproži, a za samo iskanje/prenos izpiska niso
# potrebni in samo podaljšujejo čakanje na "networkidle"/nalaganje strani.
_BLOCKED_RESOURCE_TYPES = {"stylesheet", "image", "font", "media", "other"}
_BLOCKED_URL_SNIPPETS = (
    "google-analytics.com", "googletagmanager.com", "google.com/ads",
    "doubleclick.net", "hotjar.com", "connect.facebook.net", "facebook.com/tr",
    "clarity.ms",
)

# Zagon Chromiuma pohitrimo z izklopom funkcij, ki za headless samodejno
# iskanje/prenos PDF-ja niso potrebne (razširitve, sinhronizacija, privzete
# aplikacije, prevajalnik ...) - manj stvari, ki jih mora Chromium ob zagonu
# inicializirati, pomeni hitrejši `launch()`.
CHROMIUM_LAUNCH_ARGS = [
    "--disable-gpu",
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--blink-settings=imagesEnabled=false",
    "--disable-extensions",
    "--disable-background-networking",
    "--disable-background-timer-throttling",
    "--disable-backgrounding-occluded-windows",
    "--disable-renderer-backgrounding",
    "--disable-sync",
    "--disable-default-apps",
    "--disable-translate",
    "--no-first-run",
    "--metrics-recording-only",
    "--mute-audio",
]




class EzkError(Exception):
    """Napaka med avtomatizacijo prijave/iskanja v e-ZK."""

def _require_playwright():
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
    except ImportError as e:
        raise EzkError(
            "Manjka knjižnica 'playwright' (ali je nameščena za drug Python, kot "
            "ga uporablja ta aplikacija). Namestite jo z:\n"
            "  pip install playwright --break-system-packages\n"
            "Aplikacijo nato zaženite z istim Pythonom, npr.:\n"
            "  python -m streamlit run app.py"
        ) from e


def interactive_login() -> tuple[str, dict]:
    """Odpre brskalnik (headless=False) za prijavo s SI-PASS.
    Vrne (ime_prijavljenega_uporabnika, storage_state).
    """
    _require_playwright()
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

    pw = sync_playwright().start()
    try:
        # Prijava mora potekati SAMO preko nameščenega Google Chrome (boljša
        # integracija z OS certifikati za SI-PASS) - namesto tihega prehoda
        # na ločeni Playwright Chromium ob napaki jasno sporočimo, da Chrome
        # ni bil najden.
        try:
            browser = pw.chromium.launch(headless=False, channel="chrome")
        except Exception as e:
            raise EzkError(
                "Google Chrome ni bil najden na tem računalniku (ali ga Playwright "
                "ne zazna). Prijava preko SI-PASS zahteva nameščen Google Chrome. "
                "Namestite ga z https://www.google.com/chrome/ ali poženite:\n"
                "  playwright install chrome"
            ) from e

        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # Postavi okno brskalnika v ospredje (nad glavno aplikacijo) - brez
        # tega bi se lahko odprlo v ozadju in bi uporabnik moral ročno
        # preklopiti nanj, preden lahko potrdi certifikat.
        try:
            page.bring_to_front()
        except Exception:
            pass

        page.goto(SI_PASS_LOGIN_URL, wait_until="domcontentloaded")

        # Ponovni poskus po nalaganju strani - nekateri OS/brskalniki okno ob
        # navigaciji znova potisnejo v ozadje.
        try:
            page.bring_to_front()
        except Exception:
            pass

        # Počakajmo na preusmeritev na eSodstvo SPA (potrjena prijava)
        try:
            page.wait_for_url(f"{ESODISCE_ORIGIN}/spa/**", timeout=LOGIN_TIMEOUT_MS)
            page.wait_for_load_state("networkidle", timeout=15000)
        except PWTimeout as e:
            raise EzkError(
                "Prijava ni bila potrjena v pričakovanem času ali pa ste okno zaprli. "
                "Preverite, ali imate nameščen ustrezen certifikat."
            ) from e

        # Poskusimo izluščiti napis "Prijavljeni ste kot..."
        username = "Prijavljeni ste v eSodstvo"
        try:
            page.wait_for_timeout(2000)
            js_code = """() => {
                let walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while(node = walker.nextNode()) {
                    if(node.nodeValue.includes("Prijavljeni ste kot")) {
                        return node.nodeValue.trim();
                    }
                }
                let el = document.querySelector('app-user-profile, .user-name, [class*="user"], .profile');
                if(el && el.innerText) return el.innerText.trim();
                return null;
            }"""
            res = page.evaluate(js_code)
            if res:
                username = res
        except Exception:
            pass

        state = context.storage_state()

        # Prijava je potrjena in stanje seje (piškotki) je shranjeno - brskalno
        # okno zapremo, saj ga uporabnik od tu naprej ne potrebuje več.
        try:
            context.close()
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass

        return username, state
    finally:
        try:
            pw.stop()
        except Exception:
            pass


def _find_field_locator(page, label_text: str, timeout_ms: int):
    """Poišče element vnosnega polja glede na besedilo oznake in vrne locator,
    NE da bi vanj še kaj vpisal. Poskuša več strategij, ker oznaka in polje na
    strani eSodstva nista nujno formalno povezana (brez <label for="...">).
    Vrne (locator, "input"|"combobox") ali sproži zadnjo napako.

    OPTIMIZACIJA (hitrost): prejšnja različica je ob vsaki NEUSPEŠNI strategiji
    počakala na CEL `timeout_ms` (npr. 15s), preden je poskusila naslednjo -
    ker za esodisce.si prve 1-2 strategiji (klasičen <label for="...">) skoraj
    vedno spodletita, se je to v praksi seštelo v 20-30+ sekund izgubljenega
    časa PRED vsakim uspešnim izpolnjevanjem polja. Zdaj vsako strategijo
    najprej preverimo s poceni, takojšnjim `locator.count()` (brez čakanja) -
    šele če strategija dejansko najde element v DOM-u, počakamo (kratek,
    zmanjšan timeout) še na njegovo vidnost. Ker je panel do te točke že
    prisilno odprt (glej _open_accordion_panel), je element praviloma viden
    takoj, zato kratek timeout ne ogroža zanesljivosti."""
    last_error = None
    fast_timeout_ms = min(timeout_ms, 3000)
    strategies = [
        # 1. Standardna dostopnostna povezava (če obstaja)
        ("input", lambda: page.get_by_label(label_text, exact=True)),
        ("input", lambda: page.get_by_label(label_text, exact=False)),
        ("input", lambda: page.get_by_placeholder(label_text)),
        # 2. Potrjen primer za "Katastrska občina": jQuery UI Autocomplete
        #    (class="ac_input ... ui-autocomplete-input") - iščemo prvi tak
        #    input za oznako, znotraj iste vrstice/vsebnika.
        ("combobox", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=ancestor::*[self::div or self::tr or self::li][1]"
            "//input[contains(@class,'ui-autocomplete-input') or contains(@class,'ac_input')]"
        )),
        ("combobox", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=following::input[contains(@class,'ui-autocomplete-input') or contains(@class,'ac_input')][1]"
        )),
        # 3. ARIA combobox/listbox vzorec - "Katastrska občina" ima ~2700
        #    možnih vrednosti, zato je na tovrstnih portalih skoraj vedno
        #    iskalni spustni seznam (combobox), NE navaden <input>. Element z
        #    role="combobox" je lahko sam po sebi tipkalen (contenteditable),
        #    zato ga obravnavamo posebej (glej _fill_labeled_field spodaj).
        ("combobox", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=following::*[@role='combobox' or @role='searchbox'][1]"
        )),
        ("combobox", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=ancestor::*[self::div or self::tr or self::li][1]"
            "//*[@role='combobox' or @role='searchbox'][1]"
        )),
        # 4. Besedilo oznake (kakršenkoli element - div/span/label) in nato
        #    prvi naslednji <input> v vrstnem redu DOM-a (ujema se s postavitvijo
        #    "oznaka: [prazno polje]" na dejanski strani).
        ("input", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=following::input[1]"
        )),
        # 5. Enako, a omejeno na najbližjega starša-vrstico (bolj varno, če je
        #    na strani več podobnih oznak).
        ("input", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=ancestor::*[self::div or self::tr or self::li][1]//input[1]"
        )),
        # 6. Splošen "vnosni ovoj" brez pravega <input> (npr. PrimeVue/Vuetify
        #    komponenta, kjer je dejanski <input> skrit, klikljiv pa je zunanji
        #    div s class-om, ki vsebuje "select"/"dropdown"/"autocomplete").
        ("combobox", lambda: page.get_by_text(label_text, exact=False).first.locator(
            "xpath=ancestor::*[self::div or self::tr or self::li][1]"
            "//*[contains(@class,'select') or contains(@class,'dropdown') "
            "or contains(@class,'autocomplete') or contains(@class,'combo')][1]"
        )),
    ]
    for kind, strategy in strategies:
        try:
            locator = strategy().first
            # Poceni, takojšen obstoj v DOM-u - brez tega bi vsaka spodletela
            # strategija počakala poln timeout, preden bi šli na naslednjo.
            try:
                if locator.count() == 0:
                    continue
            except Exception:
                continue
            locator.wait_for(state="visible", timeout=fast_timeout_ms)
            return locator, kind
        except Exception as e:
            last_error = e
            continue
    if last_error:
        raise last_error
    raise RuntimeError(f"Polje '{label_text}' ni bilo najdeno.")


def _fill_labeled_field(page, label_text: str, value: str, timeout_ms: int = 5_000):
    """Izpolni vnosno polje ekstremno hitro prek direktnega DOM vnosa, da se izognemo napakam 
    zaradi prekrivanja elementov (ker je CSS izklopljen)."""
    locator, kind = _find_field_locator(page, label_text, timeout_ms)
    
    safe_value = str(value).replace("'", "\'")
    
    # Namesto zanašanja na vmesnik in padajoče menije, vrednost vpišemo naravnost v kodo elementa.
    locator.first.evaluate(f"el => el.value = '{safe_value}'")
    locator.first.evaluate("el => el.dispatchEvent(new Event('input', { bubbles: true }))")
    locator.first.evaluate("el => el.dispatchEvent(new Event('change', { bubbles: true }))")
    locator.first.evaluate("el => el.dispatchEvent(new Event('blur', { bubbles: true }))")
    
    return True


def _dump_debug_snapshot(page, prefix: str, extra_locator=None) -> str:
    """Ob napaki shrani zaslonsko sliko in HTML trenutne strani, da je iskanje
    mogoče vizualno preveriti. Če je podan extra_locator (npr. okolica oznake,
    ki je povzročila napako), shrani dodatno tudi njegov outerHTML v ločeno
    datoteko - to je največkrat najhitrejši način, da se ugotovi, ali se je
    razporeditev polj na strani spremenila (npr. iz <input> v combobox).
    Vrne opis shranjenih poti (ali prazen niz, če shranjevanje ni uspelo)."""
    try:
        out_dir = Path(tempfile.mkdtemp(prefix="ezk_debug_"))
        png_path = out_dir / f"{prefix}.png"
        html_path = out_dir / f"{prefix}.html"
        page.screenshot(path=str(png_path), full_page=True)
        html_path.write_text(page.content(), encoding="utf-8")
        msg = f" Zaslonska slika in HTML strani sta shranjena v: {out_dir}"
        if extra_locator is not None:
            try:
                snippet_path = out_dir / f"{prefix}_element.html"
                outer = extra_locator.first.evaluate("el => el.outerHTML")
                snippet_path.write_text(outer or "", encoding="utf-8")
                msg += f" (glej tudi {snippet_path.name} za HTML okolice oznake)"
            except Exception:
                pass
        return msg
    except Exception:
        return ""


def _open_accordion_panel(page, header_text: str, timeout_ms: int = 20_000) -> bool:
    """Odpre panel jQuery UI accordion widgeta (npr. "03-001 - Redni izpis iz
    zemljiške knjige") na PROGRAMSKI način, namesto z navadnim klikom.

    Portal esodisce.si NI Nuxt/Vue SPA (kot so napačno predvidevali prejšnji
    komentarji v tej datoteki), temveč klasična, strežniško izrisana stran z
    jQuery UI accordion widgetom (`<div id="accordion" class="ui-accordion...">`).
    Vsi paneli so v HTML-ju izrisani vnaprej in samo skriti (`display: none`),
    dokler jih accordion "ne odpre".

    Prejšnja implementacija je samo kliknila na glavo panela in upala, da se
    bo jQuery UI drsna animacija ("slide") uspešno zaključila. Težava: jQuery
    UI ob kliku `aria-expanded="true"` na glavi nastavi TAKOJ (sinhrono), še
    preden animacija sploh steče - če se nato animacija prekine (JS napaka,
    prehiter naslednji korak ...), `aria-expanded` ostane "true", vsebina pa
    ostane `display: none`. `aria-expanded` je torej sam po sebi nezanesljiv
    pokazatelj, da je panel res viden.

    Ta funkcija namesto tega:
      1. Najprej preveri, ali je panel DEJANSKO viden (računano preko
         `getComputedStyle`, ne samo `aria-expanded`) - če je, ni česa delati.
      2. Če ni viden, pokliče neposredno jQuery UI accordion widget API
         (`.accordion('option', 'active', index)`) z izklopljeno animacijo -
         to postavi stanje panela TAKOJ, brez tveganja prekinjene animacije.
      3. Če jQuery UI widget iz kakršnegakoli razloga ni dosegljiv/ne uspe, kot
         zadnji poskus prisilno razkrije vsebino neposredno preko DOM-a (ker je
         itak že v celoti izrisana, samo skrita).
      4. Na koncu vedno znova preveri DEJANSKO vidnost vsebine, preden vrne
         True/False - klicatelj se torej ne rabi zanašati samo na to, da klic
         ni vrgel izjeme.
    """
    header = page.get_by_text(header_text, exact=False).first
    header.wait_for(state="visible", timeout=timeout_ms)

    def _panel_state():
        return header.evaluate(
            """(h) => {
                const hdr = h.closest('h3, .ui-accordion-header') || h;
                let panel = hdr.nextElementSibling;
                if (!panel) {
                    const ctrl = hdr.getAttribute('aria-controls');
                    if (ctrl) panel = document.getElementById(ctrl);
                }
                const visible = !!panel && panel.offsetParent !== null
                    && getComputedStyle(panel).display !== 'none';
                return {
                    expanded: hdr.getAttribute('aria-expanded') === 'true',
                    visible,
                    hasPanel: !!panel,
                };
            }"""
        )

    if _panel_state().get("visible"):
        return True

    # Poskus 1: neposreden klic jQuery UI accordion widget API-ja (brez
    # animacije - takojšnja sprememba stanja, brez tveganja prekinitve).
    header.evaluate(
        """(h) => {
            if (!window.jQuery) return {ok: false, reason: 'no-jquery'};
            const $ = window.jQuery;
            const $hdr = $(h).closest('h3, .ui-accordion-header');
            const $acc = $hdr.closest('.ui-accordion');
            if (!$acc.length || typeof $acc.accordion !== 'function') {
                return {ok: false, reason: 'no-accordion-widget'};
            }
            try {
                $acc.accordion('option', 'animate', false);
                const $headers = $acc.find('> h3, > .ui-accordion-header');
                const idx = $headers.index($hdr);
                if (idx < 0) return {ok: false, reason: 'header-not-in-list'};
                $acc.accordion('option', 'active', idx);
                return {ok: true, index: idx};
            } catch (e) {
                return {ok: false, reason: String(e)};
            }
        }"""
    )

    if _panel_state().get("visible"):
        return True

    # Poskus 2: zadnji poskus - panel je itak v celoti izrisan v HTML-ju (samo
    # skrit), zato ga prisilno razkrijemo neposredno preko DOM-a. To ne obnovi
    # "lepega" izgleda puščice/animacije, a zagotovi, da so polja znotraj
    # panela vidna in dosegljiva za izpolnjevanje.
    header.evaluate(
        """(h) => {
            const hdr = h.closest('h3, .ui-accordion-header') || h;
            let panel = hdr.nextElementSibling;
            if (!panel) {
                const ctrl = hdr.getAttribute('aria-controls');
                if (ctrl) panel = document.getElementById(ctrl);
            }
            hdr.setAttribute('aria-expanded', 'true');
            hdr.classList.add('ui-accordion-header-active', 'ui-state-active');
            if (panel) {
                panel.style.display = 'block';
                panel.removeAttribute('aria-hidden');
                panel.classList.add('ui-accordion-content-active');
            }
        }"""
    )

    return bool(_panel_state().get("visible"))


def fetch_redni_izpis_pdf_with_state(
    state: dict,
    katastrska_obcina: str,
    parcelna_stevilka: str,
    headless: bool = True,
) -> bytes:
    """Z uporabo prejšnjega stanja prijave (piškotkov) izvede iskanje in prenese PDF."""
    _require_playwright()
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

    pw = sync_playwright().start()

    try:
        # OPTIMIZACIJA: skupni, pohitreni nabor Chromium zastavic (glej
        # CHROMIUM_LAUNCH_ARGS zgoraj) - manj notranjih Chromium podsistemov
        # za inicializacijo pomeni hitrejši `launch()`.
        browser = pw.chromium.launch(headless=headless, args=CHROMIUM_LAUNCH_ARGS)
        context = browser.new_context(
            storage_state=state,
            accept_downloads=True,
            # OPTIMIZACIJA: service workerji za samo iskanje/prenos PDF-ja
            # niso potrebni in lahko sprožijo dodatne ozadenjske zahteve.
            service_workers="block",
        )
        page = context.new_page()
        # Stran NI Nuxt/Vue SPA, temveč klasična, strežniško izrisana stran z
        # jQuery UI accordion widgetom. Ker je nabor blokiranih virov širši
        # (glej OPTIMIZACIJA 1 spodaj), je dejansko nalaganje bistveno hitrejše
        # kot prej, zato lahko privzeti timeout znižamo (a še vedno pustimo
        # dovolj varnostne rezerve za počasnejšo povezavo).
        page.set_default_timeout(25_000)

        # ---------------- OPTIMIZACIJA 1: blokiranje nepotrebnih virov -----
        # Poleg CSS/slik/pisav/medijev (kot prej) zdaj blokiramo tudi znane
        # domene za analitiko/sledenje (glej _BLOCKED_URL_SNIPPETS) - te za
        # samo iskanje/prenos izpiska niso potrebne, klic pa je poceni
        # (preverjanje niza), zato ne upočasni "prepustnih" zahtev.
        def intercept_route(route):
            req = route.request
            if req.resource_type in _BLOCKED_RESOURCE_TYPES:
                route.abort()
                return
            url = req.url
            if any(snippet in url for snippet in _BLOCKED_URL_SNIPPETS):
                route.abort()
                return
            route.continue_()
        page.route("**/*", intercept_route)
        # --------------------------------------------------------------------

        # 1. Povezava na ezk access endpoint za potrditev seje - zgolj
        # "sprožimo" zahtevo (fire-and-forget, ovita v try/except), zato
        # čakamo samo na "commit" (prejet odgovor) in s kratkim timeoutom.
        try:
            page.goto(ACCESS_URL, wait_until="commit", timeout=6000)
        except Exception:
            pass

        # 2. Povezava na iskalnik - OPTIMIZACIJA: "commit" namesto
        # "domcontentloaded" (ne čakamo na razčlenitev celotnega DOM-a preden
        # gremo naprej), dejansko pripravljenost strani pa ugotovimo
        # neposredno s ciljanim čakanjem na vstopno točko (spodaj) - to je
        # hitreje, ker ne čakamo dvakrat zaporedno na podoben pogoj.
        page.goto(SEARCH_URL, wait_until="commit", timeout=30000)
        try:
            page.wait_for_selector("text=03-001 - Redni izpis iz zemljiške knjige", state="attached", timeout=15000)
        except PWTimeout:
            pass

        # Portal je klasična stran z jQuery UI accordion widgetom - vsi paneli
        # so v HTML-ju izrisani vnaprej, samo skriti (display: none), dokler
        # jih accordion ne odpre. Panel odpremo na PROGRAMSKI način (glej
        # _open_accordion_panel), ne z navadnim klikom: izkazalo se je, da klik
        # prek jQuery UI drsne animacije ni bil zanesljiv - aria-expanded se je
        # postavil na "true", vsebina pa je ostala display: none (prekinjena/
        # spodletela animacija). _open_accordion_panel na koncu vedno preveri
        # DEJANSKO (computed-style) vidnost vsebine, ne samo aria-expanded.
        try:
            panel_opened = _open_accordion_panel(
                page, "03-001 - Redni izpis iz zemljiške knjige", timeout_ms=20_000
            )
        except PWTimeout:
            panel_opened = False

        if not panel_opened:
            snap = _dump_debug_snapshot(page, "napaka_odpiranje_panela")
            raise EzkError(
                "Panela 'Redni izpis iz zemljiške knjige' ni bilo mogoče odpreti "
                "(niti prek jQuery UI accordion widget API-ja, niti prek "
                f"neposrednega DOM zapisa).{snap}"
            )

        try:
            page.get_by_label("Način vnosa nepremičnin").select_option(label="po ID znaku", force=True)
        except Exception:
            pass
        try:
            page.get_by_label("Tip nepremičnine").select_option(label="zemljiška parcela", force=True)
        except Exception:
            pass

        try:
            _fill_labeled_field(page, "Katastrska občina", katastrska_obcina)
        except Exception as e:
            try:
                label_area = page.get_by_text("Katastrska občina", exact=False).first.locator(
                    "xpath=ancestor::*[self::div or self::tr or self::li][1]"
                )
            except Exception:
                label_area = None
            snap = _dump_debug_snapshot(page, "napaka_katastrska_obcina", extra_locator=label_area)
            raise EzkError(
                "Ni bilo mogoče najti/izpolniti polja 'Katastrska občina' na strani "
                "eSodstva (stran se je morda spremenila, polje je morda iskalni spustni "
                "seznam namesto navadnega polja, ali pa se panel ni pravočasno "
                f"razširil).{snap}\n\nIzvirna napaka: {e}"
            ) from e

        try:
            _fill_labeled_field(page, "Parcelna številka", parcelna_stevilka)
        except Exception as e:
            snap = _dump_debug_snapshot(page, "napaka_parcelna_stevilka")
            raise EzkError(
                f"Ni bilo mogoče najti/izpolniti polja 'Parcelna številka'.{snap}"
            ) from e

        submit_btn = page.get_by_role("button", name="Prikaži v pdf obliki")

        try:
            with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl_info:
                submit_btn.click(force=True)
            download = dl_info.value
            tmp_path = Path(tempfile.mkdtemp(prefix="ezk_pdf_")) / "izpisek.pdf"
            download.save_as(str(tmp_path))
            return tmp_path.read_bytes()
        except PWTimeout:
            pass

        try:
            with page.context.expect_page(timeout=DOWNLOAD_TIMEOUT_MS) as new_page_info:
                submit_btn.click(force=True)
            new_page = new_page_info.value
            new_page.wait_for_load_state("domcontentloaded")
            resp = new_page.request.get(new_page.url)
            pdf_bytes = resp.body()
            new_page.close()
            if not pdf_bytes.startswith(b"%PDF"):
                raise EzkError("Stran se je odprla, a vsebina ne izgleda kot veljaven PDF. Preverite podatke.")
            return pdf_bytes
        except PWTimeout as e:
            raise EzkError(
                "Klik na 'Prikaži v pdf obliki' ni sprožil prenosa. Preverite, "
                "ali so vneseni podatki točni."
            ) from e
    finally:
        try:
            context.close()
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass
        try:
            pw.stop()
        except Exception:
            pass
