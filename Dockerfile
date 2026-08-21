# ============================================================
#  e-SlužnostiMB - Docker slika za spletno gostovanje
#
#  Zgradi in zaženi:
#    docker build -t e-sluznostimb .
#    docker run -p 8501:8501 --env-file .env e-sluznostimb
#
#  Nato je aplikacija dostopna na http://<naslov-streznika>:8501
#  Za pravo domeno (https://mojadomena.si) postavite pred to sliko
#  reverse proxy (nginx/Caddy/Traefik) - glej README_DEPLOY_WEB.md.
# ============================================================
FROM python:3.12-slim

# Sistemske odvisnosti:
# - libreoffice + poppler-utils: za vizualno primerjavo dokumentov pred/po
# - odvisnosti za Playwright/Chromium: za samodejno prijavo v eZK (neobvezno)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    wget \
    gnupg \
    fonts-liberation \
    libnss3 \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    libgbm1 \
    libasound2 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Chromium za Playwright (samodejna prijava v eZK) - lahko izpustite to
# vrstico, če te funkcije ne potrebujete (zmanjša velikost slike).
RUN python -m playwright install --with-deps chromium

COPY . .

EXPOSE 8501

HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.port=8501", \
            "--server.address=0.0.0.0", \
            "--server.headless=true"]
