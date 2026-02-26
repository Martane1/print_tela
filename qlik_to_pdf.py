import os
import re
from datetime import datetime
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ====== CONFIGURE AQUI ======
QLIK_URL = r"https://acesso.sigaer.intraer:5432/mashup/sense/app/85e81e92-7767-4397-896a-3c78068598b8/sheet/c7aa1aef-f126-4b57-9479-d5911d96add3/state/analysis"
OUTPUT_DIR = r"C:\QlikPrint\pdfs"
TMP_DIR = r"C:\QlikPrint\tmp"
WAIT_MAX_MS = 300_000
# Ordem igual ao que aparece no menu do seu PDF (ajuste se tiver mais/menos)
MENU_OMS = [
    "PAGINA INICIAL",
    "AFA",
    "EPCAR",
    "EEAR",
    "CIAAR",
    "IEAD",
    "UNIFA",
    "ECEMAR",
    "EAOAR",
    "DIRENS",
    "ASSISTENCIAIS",
]
# ============================

def safe(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:120] if len(s) > 120 else s

def wait_qlik(page, extra_ms=0):
    # Qlik é SPA; networkidle nem sempre estabiliza. Mistura de esperas.
    page.wait_for_timeout(4000 + extra_ms)
    try:
        page.wait_for_load_state("networkidle", timeout=12000)
    except PWTimeout:
        pass
    page.wait_for_timeout(5000 + extra_ms)

def click_menu_item(page, name: str):
    # tenta clicar como botão (melhor), senão fallback por texto
    loc = page.get_by_role("button", name=name)
    if loc.count() > 0:
        loc.first.click()
        return True

    # fallback: qualquer elemento clicável contendo o texto
    loc2 = page.locator(f"text={name}")
    if loc2.count() > 0:
        loc2.first.click()
        return True

    return False

def click_if_exists(page, label: str) -> bool:
    # “Detalhar”, “Voltar” etc.
    loc = page.get_by_role("button", name=label)
    if loc.count() > 0:
        loc.first.click()
        return True

    loc2 = page.locator(f"text={label}")
    if loc2.count() > 0:
        loc2.first.click()
        return True

    return False

def screenshot_page(page, out_png: str):
    page.screenshot(path=out_png, full_page=False)

def build_pdf(png_paths, out_pdf):
    imgs = [Image.open(p).convert("RGB") for p in png_paths]
    if not imgs:
        raise RuntimeError("Nenhuma imagem gerada.")
    first, rest = imgs[0], imgs[1:]
    first.save(out_pdf, save_all=True, append_images=rest)

def main():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    Path(TMP_DIR).mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_pdf = os.path.join(OUTPUT_DIR, f"e-GovEns_{ts}.pdf")

    # limpa tmp antigo
    for f in Path(TMP_DIR).glob("page_*.png"):
        try: f.unlink()
        except: pass

    shots = []
    idx = 1

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,       # deixe False (Qlik mais estável); depois dá pra testar True
            channel="chrome"      # usa Chrome instalado (sem baixar)
        )
        context = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
            device_scale_factor=1
        )
        page = context.new_page()

        page.goto(QLIK_URL, wait_until="domcontentloaded", timeout=WAIT_MAX_MS)
        wait_qlik(page, extra_ms=4000)

        for om in MENU_OMS:
            ok = click_menu_item(page, om)
            if not ok:
                print(f"[AVISO] Não achei no menu: {om} (vou pular).")
                continue

            # espera carregar a tela da OM
            wait_qlik(page, extra_ms=6000)

            # captura a tela “principal” da OM
            png = os.path.join(TMP_DIR, f"page_{idx:03d}_{safe(om)}.png")
            screenshot_page(page, png)
            shots.append(png)
            print(f"[OK] Capturada: {om}")
            idx += 1

            # Se existir “Detalhar”, entra e captura 1x (ou mais, se houver vários cliques)
            # Loop: Detalhar -> print -> Voltar -> (tenta outro Detalhar)
            # Obs: se tiver vários “Detalhar” diferentes, isso pode pegar só o primeiro.
            for _ in range(6):  # trava de segurança
                if not click_if_exists(page, "Detalhar"):
                    break
                wait_qlik(page, extra_ms=7000)

                png = os.path.join(TMP_DIR, f"page_{idx:03d}_{safe(om)}_DETALHAR.png")
                screenshot_page(page, png)
                shots.append(png)
                print(f"[OK] Capturada: {om} -> Detalhar")
                idx += 1

                # voltar
                if click_if_exists(page, "Voltar"):
                    wait_qlik(page, extra_ms=5000)
                else:
                    # às vezes o “Voltar” é seta/ícone; se não achar, sai do loop
                    break

        context.close()
        browser.close()

    build_pdf(shots, out_pdf)
    print(f"\nOK - PDF no mesmo estilo gerado em: {out_pdf}")
    print(f"Imagens temporárias em: {TMP_DIR}")

if __name__ == "__main__":
    main()