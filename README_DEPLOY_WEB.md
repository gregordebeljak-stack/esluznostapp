# Kako objaviti e-SlužnostiMB kot spletno aplikacijo

Ta mapa je urejena tako, da jo lahko **neposredno objaviš na spletu** - brez
gradnje `.exe`, brez namestitve na vsak posamezen računalnik. Uporabniki
odprejo samo povezavo (URL) v brskalniku.

Na voljo sta dve poti - izberi eno glede na potrebe:

| | Streamlit Community Cloud | Lastna domena (Docker) |
|---|---|---|
| Cena | brezplačno | strošek strežnika/gostovanja |
| Zahtevnost | najlažje (par klikov) | zahteva osnovno znanje strežnikov |
| Lastna domena (npr. sluznosti.elektro-maribor.si) | ne (samo *.streamlit.app) | da |
| Samodejna prijava v eZK (Playwright/Chromium) | **ne deluje** (glej opombo spodaj) | deluje |
| Nadzor nad podatki | gostuje Streamlit (javno oblačje) | gostuje na tvojem strežniku |

---

## POT A - Streamlit Community Cloud (brezplačno, najhitreje)

### 1. Naloži kodo na GitHub
```bash
cd pot/do/te/mape
git init
git add .
git commit -m "Prva objava"
git branch -M main
git remote add origin https://github.com/<tvoj-uporabnik>/<ime-repozitorija>.git
git push -u origin main
```
`.env` ni vključen (je v `.gitignore`), zato tvoj API ključ ne bo pomotoma
pristal na GitHubu.

### 2. Ustvari aplikacijo na Streamlit Cloud
1. Pojdi na https://share.streamlit.io in se prijavi z GitHub računom.
2. Klikni **"New app"**, izberi svoj repozitorij, vejo `main` in glavno
   datoteko `app.py`.
3. Pod **"Settings → Secrets"** vnesi API ključ v formatu TOML (glej primer v
   `.streamlit/secrets.toml.example`):
   ```toml
   NVIDIA_API_KEY = "vaš-pravi-ključ"
   ```
4. Klikni **"Deploy"**. Prvi zagon traja nekaj minut (namesti se
   `requirements.txt` in `packages.txt` - slednji poskrbi za LibreOffice in
   poppler, potrebna za vizualno primerjavo dokumentov).

Po končani objavi dobiš javno povezavo oblike:
```
https://<ime-aplikacije>.streamlit.app
```
To povezavo pošlješ sodelavcem - odprejo jo v brskalniku, brez nameščanja.

### Omejitev na tej poti
Streamlit Community Cloud ne omogoča namestitve Chromium brskalnika, ki ga
`ezk_engine.py` uporablja za **samodejno prijavo v eZK** (Playwright). Ta
konkretna funkcija na tej poti ne bo delovala - aplikacija ob poskusu izpiše
jasno sporočilo, vse ostalo (ročno nalaganje PDF izpiskov, izpolnjevanje
pogodb, primerjava dokumentov) pa deluje normalno. Če to funkcijo
potrebuješ, uporabi POT B spodaj.

---

## POT B - lastna domena prek Docker (polna funkcionalnost)

Ta pot zahteva svoj strežnik (VPS, npr. Hetzner, DigitalOcean, ali strežnik
znotraj Elektro Maribor) z nameščenim Dockerjem, na katerem ima aplikacija
polni nadzor - vključno s Playwright/Chromium za samodejno prijavo v eZK.

### 1. Pripravi `.env`
```bash
cp .env.example .env
# uredi .env in vpiši svoj pravi API ključ
```

### 2. Zgradi in zaženi (na strežniku, kjer je ta mapa)
```bash
docker compose up -d --build
```
Aplikacija je zdaj dostopna na `http://<naslov-streznika>:8501`.

### 3. Poveži z lastno domeno (HTTPS)
Pred aplikacijo postavi reverse proxy, ki poskrbi za HTTPS in preusmeri
promet z domene na vrata 8501. Najlažje s **Caddy** (samodejno pridobi
brezplačen HTTPS certifikat):

```bash
# Caddyfile
sluznosti.elektro-maribor.si {
    reverse_proxy localhost:8501
}
```
```bash
caddy run
```

Po tem koraku je aplikacija dostopna na `https://sluznosti.elektro-maribor.si`
brez dodatnih nastavitev certifikatov.

### Posodobitev aplikacije kasneje
```bash
git pull            # če uporabljaš git za sledenje kodi
docker compose up -d --build
```

---

## Kaj se je spremenilo v strukturi mape

Za spletno objavo sem odstranil datoteke, ki so bile namenjene samo Windows
`.exe` gradnji (niso potrebne za splet in bi zmedle):
- `app.spec`, `e-SlužnostiMB.spec` (PyInstaller)
- `run_app.py`, `run_app.pyw`, `launch_app.vbs` (Windows zagonske skripte)
- `build_exe.bat`, `build_installer.bat`, `installer.iss` (Windows build/installer)

Dodal sem:
- `Dockerfile` - slika za samostojno gostovanje (POT B)
- `docker-compose.yml` - enostaven zagon in samodejni ponovni zagon ob padcu
- ta navodila (`README_DEPLOY_WEB.md`)

Vse ostalo (`app.py`, `*_engine.py`, `templates/`, `assets/`, `requirements.txt`,
`packages.txt`, `.streamlit/`) je nespremenjeno - to je ista aplikacija, samo
brez desktop ovojnice.

Če kasneje vseeno želiš tudi Windows `.exe` različico (npr. za delo brez
internetne povezave), mi povej - tisto mapo (`build_pkg`) še vedno imam
pripravljeno ločeno.
