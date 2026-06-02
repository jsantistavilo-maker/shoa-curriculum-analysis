"""
Script para descargar el logo oficial del SHOA.
Ejecutar una vez antes de lanzar la app:  python setup_assets.py
"""

import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import requests
from bs4 import BeautifulSoup

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH  = ASSETS_DIR / "logo_shoa.png"

DIRECT_URLS = [
    "https://www.shoa.cl/php/images/logo_shoa.png",
    "https://www.shoa.cl/images/logo_shoa.png",
    "https://www.shoa.cl/php/images/logo.png",
    "https://www.shoa.cl/php/images/logo_shoa.jpg",
    "https://www.shoa.cl/images/logo.png",
]

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

SVG_PLACEHOLDER = """\
<svg width="200" height="80" xmlns="http://www.w3.org/2000/svg">
  <rect width="200" height="80" fill="#003366" rx="6"/>
  <text x="100" y="28" font-family="Arial" font-size="22"
        font-weight="bold" fill="#C8A84B" text-anchor="middle">⚓ SHOA</text>
  <text x="100" y="48" font-family="Arial" font-size="8"
        fill="white" text-anchor="middle">SERVICIO HIDROGRÁFICO Y</text>
  <text x="100" y="62" font-family="Arial" font-size="8"
        fill="white" text-anchor="middle">OCEANOGRÁFICO - ARMADA DE CHILE</text>
</svg>"""


def _guardar_logo(data: bytes) -> bool:
    if len(data) < 1000:
        return False
    sig = data[:4]
    if not (sig[:4] == b"\x89PNG" or sig[:3] == b"\xff\xd8\xff"
            or sig[:4] in (b"GIF8", b"RIFF") or b"<svg" in data[:200].lower()):
        return False
    ASSETS_DIR.mkdir(exist_ok=True)
    LOGO_PATH.write_bytes(data)
    return True


def descargar_logo_shoa() -> bool:
    # 1. Intentar URLs directas
    for url in DIRECT_URLS:
        try:
            r = requests.get(url, headers=HEADERS, timeout=8)
            if r.status_code == 200 and _guardar_logo(r.content):
                print(f"✅ Logo descargado desde: {url}")
                return True
        except Exception:
            pass

    # 2. Scraping de la página principal
    try:
        r = requests.get("https://www.shoa.cl", headers=HEADERS, timeout=10)
        soup = BeautifulSoup(r.content, "html.parser")
        for img in soup.find_all("img"):
            src = img.get("src", "")
            alt = img.get("alt", "").lower()
            if "logo" in src.lower() or "logo" in alt or "shoa" in alt:
                url_logo = (src if src.startswith("http")
                            else "https://www.shoa.cl/" + src.lstrip("/"))
                img_data = requests.get(url_logo, headers=HEADERS, timeout=8).content
                if _guardar_logo(img_data):
                    print(f"✅ Logo descargado desde: {url_logo}")
                    return True
    except Exception as e:
        print(f"   Scraping falló: {e}")

    return False


def crear_placeholder_svg() -> None:
    ASSETS_DIR.mkdir(exist_ok=True)
    (ASSETS_DIR / "logo_shoa.svg").write_text(SVG_PLACEHOLDER, encoding="utf-8")
    print("⚠️  No se pudo obtener logo, usando placeholder institucional SVG")
    print(f"   Guardado en: {ASSETS_DIR / 'logo_shoa.svg'}")


def main():
    print("🔍 Buscando logo oficial del SHOA ...")
    if descargar_logo_shoa():
        print(f"📁 Guardado en: {LOGO_PATH}")
    else:
        crear_placeholder_svg()


if __name__ == "__main__":
    main()
