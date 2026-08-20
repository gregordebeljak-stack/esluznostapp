# Nadzorna plošča za avtomatizacijo dokumentov

Streamlit aplikacija, ki iz PDF izpiska iz slovenske zemljiške knjige (eZK)
izlušči pet ključnih podatkov - **ime in priimek, naslov, katastrsko občino,
parcelno številko in delež** - ter jih samodejno zamenja v ustreznih rdeče
obarvanih poljih .docx predloge. Prikaže tudi vizualno primerjavo dokumenta
pred in po izpolnitvi.

## 1. Namestitev Python odvisnosti

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. (Priporočeno) Vizualna primerjava pred/po

Za pravi vizualni predogled dokumenta (z barvami, tabelami, pisavami) uporablja
aplikacija LibreOffice + poppler-utils. Brez njiju aplikacija še vedno deluje
(izpolnjevanje, izvoz), a brez slikovne primerjave strani.

**Windows/Mac:** naložite LibreOffice s https://www.libreoffice.org/download/
in poppler (npr. `choco install poppler` na Windows ali `brew install poppler`
na Mac).

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install libreoffice poppler-utils
```

## 3. Nastavitev API ključa (BREZ ročnega vnašanja vsakič)

Aplikacija uporablja [NVIDIA API](https://build.nvidia.com) (NVIDIA NIM) - en
sam ključ, gostovan (cloud) dostop do velikega izbora modelov (Meta Llama,
NVIDIA Nemotron, DeepSeek, Mistral, Google Gemma ...). Ključ ustvarite na
https://build.nvidia.com - odprite poljuben model in kliknite "Get API Key"
(ključ se začne z "nvapi-").

**Opomba glede naslova API-ja (base URL):** privzeto aplikacija uporablja
`https://integrate.api.nvidia.com/v1`. Če uporabljate drugačen naslov (npr.
lastno gostovan NIM strežnik), ga vpišite v polje "Naslov NVIDIA API (base
URL)" v nadzorni plošči ali nastavite `NVIDIA_BASE_URL` v `.env` (glej spodaj).

**Ključa NIKOLI ne vpisujte neposredno v `app.py` ali katerokoli datoteko, ki
jo delite naprej ali nalagate na GitHub** - vsakdo z dostopom do te datoteke
bi ga lahko zlorabil na vaš račun.

Namesto tega:

1. Kopirajte `.env.example` v novo datoteko z imenom `.env` (v isti mapi).
2. V `.env` vpišite svoj pravi ključ:
   ```
   NVIDIA_API_KEY=...vaš_pravi_ključ...
   ```
3. Shranite. Aplikacija ga ob zagonu samodejno prebere - polja za ključ v
   vmesniku ne boste več videli (namesto tega bo zeleno obvestilo "API ključ
   najden").

Datoteka `.env` je v `.gitignore`, zato se po nesreči ne bo znašla v Gitu.
**Nikoli je ne pošiljajte naprej, ne nalagajte v skupne mape in ne delite v
klepetu/e-pošti.**

## 4. Zagon

```bash
streamlit run app.py
```

Aplikacija se odpre na `http://localhost:8501`.

## 5. Objava na GitHub

```bash
git init
git add .
git commit -m "Prva objava"
git branch -M main
git remote add origin https://github.com/<uporabnik>/<ime-repozitorija>.git
git push -u origin main
```

Datoteka `.env` (s pravim API ključem) je zaradi `.gitignore` samodejno
izključena iz Gita - preverite z `git status`, da se res ne pojavi v seznamu
sprememb, preden potisnete kodo na GitHub.

## 6. Namestitev (deploy) na Streamlit Community Cloud

1. Pojdite na https://share.streamlit.io in se prijavite z GitHub računom.
2. Kliknite "New app" in izberite svoj repozitorij, vejo (`main`) ter glavno
   datoteko `app.py`.
3. Pred zagonom (ali kasneje pod "Settings → Secrets") vnesite API ključ v
   polje **Secrets** v formatu TOML - glej primer v
   `.streamlit/secrets.toml.example`:
   ```toml
   NVIDIA_API_KEY = "vaš-pravi-ključ"
   ```
   Streamlit vrednosti iz Secrets samodejno izpostavi tudi kot okoljske
   spremenljivke, zato jih aplikacija prebere enako kot lokalni `.env`.
4. Kliknite "Deploy". Prvi zagon lahko traja nekaj minut, ker Streamlit
   namesti odvisnosti iz `requirements.txt` in (za vizualno primerjavo
   pred/po) sistemske pakete iz `packages.txt` (LibreOffice + poppler-utils).
5. Za posodobitev aplikacije preprosto potisnite (`git push`) spremembe na
   `main` - Streamlit Cloud jo samodejno znova zažene.

**Opomba:** LibreOffice na Streamlit Cloud poveča čas prvega zagona in porabo
pomnilnika. Če vizualna primerjava pred/po ni nujna, lahko datoteko
`packages.txt` pobrišete - aplikacija bo še vedno delovala (izpolnjevanje,
izvoz), le brez slikovnega predogleda strani (uporabila bo osnovni HTML
predogled prek knjižnice `mammoth`).

## Kako deluje

1. **Naloži predlogo (.docx)** – sistem preišče XML strukturo dokumenta in
   najde vse dele besedila, obarvane rdeče (`#FF0000` ali `#C00000`).
2. **Naloži PDF izpisek(-e) iz zemljiške knjige.** Če ima lastnik/lastnica v
   lasti VEČ parcel, naložite ločen PDF izpisek za vsako - gumb za nalaganje
   sprejme več datotek hkrati (ali jih dodajate eno za drugo).
3. **Klikni "Izvleci iz PDF-jev in zamenjaj v .docx"** – NVIDIA API (privzeto Llama 3.3 70B Instruct) prebere vsak
   izpisek posebej. Podatki PRVEGA naloženega PDF-ja se uporabijo za
   izpolnitev imena/priimka/naslova ter prve vrstice v tabelah parcel. Za
   VSAK NADALJNJI PDF sistem samodejno **podvoji vrstico** v vseh tabelah, ki
   vsebujejo podatke o parceli (parc. št. / k.o. / delež), in jo zapolni s
   podatki tega PDF-ja - oblika (obrobe, pisave, širine stolpcev) ostane
   identična, ker gre za dobesedno kopijo vrstice.
4. **Preglej ujemanja** v srednjem stolpcu - urejate lahko podatke primarne
   parcele; dodatne vrstice so že zapisane neposredno v dokument (za njihove
   ročne popravke odprite izvoženo .docx datoteko v Wordu).
5. **Ustvari vizualno primerjavo** in **izvozi** kot običajno.

### Primer: lastnik ima 3 parcele

Naložite `parcela1.pdf`, `parcela2.pdf`, `parcela3.pdf` (vse za istega
lastnika). Aplikacija bo v tabeli "parc. št. | k.o. | delež | opomba" ustvarila
3 vrstice namesto ene, vsako s pravimi podatki. Če se ime lastnika med PDF-ji
razlikuje, vas aplikacija opozori (možna napaka pri izbiri datotek).

## Znane omejitve

- Pri solastništvu z več lastniki NA ISTI parceli izbrani model vrne podatke prvega
  navedenega lastnika; ostale je treba dopolniti ročno.
- Podvajanje vrstic za več parcel deluje na TABELAH (strukturirani podatki).
  Prosto besedilo v odstavkih (npr. stavek "...na zemljišču s parc. št.: X,
  k.o. Y izrecno dovoljuje...") trenutno omenja le PRVO/primarno parcelo -
  če pogodba zajema več parcel, ta stavek po potrebi ročno dopolnite.
- Prepoznavanje polj temelji na kontekstu (ključne besede v okolici) in
  obliki vrednosti (regex). Pri zelo nenavadno oblikovanih predlogah je
  smiselno preveriti ujemanja v srednjem stolpcu, preden izvozite dokument.
- Zaznavanje rdečih placeholderjev temelji na neposredni barvi pisave na
  nivoju run-a; barve, podedovane izključno iz Wordovega sloga, niso zaznane.
- Vizualna primerjava pred/po zahteva LibreOffice + poppler-utils (glej
  točko 2 zgoraj); brez njiju se prikaže samo osnovni HTML predogled.
- Gumb "Izvleci in zamenjaj" ob vsakem kliku dokument znova sestavi iz
  izvirne predloge (da se vrstice pri ponovnem kliku ne podvajajo) - ročne
  popravke drugih polj zato naredite ŠELE po tem koraku, ne pred njim.
