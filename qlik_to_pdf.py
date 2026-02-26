import os
import re
import unicodedata
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

# Termos usados para achar cards de drill-down por OM.
# Mantém variações para tolerar acentos/grafia diferente no dashboard.
DETAIL_TERMS_BY_OM = {
    "AFA": [
        "Formação Aviadores",
        "Formação Intendência",
        "Formação Intendências",
        "Formação Infantaria",
        "Formação Infante",
        "Aeronave / Simulador",
        "Aeronave/Simulador",
        "Instrução T25",
        "Instrução T27",
        "Esforço Aéreo",
    ],
    "ASSISTENCIAIS": [
        "Clique para detalhar",
    ],
}

GENERIC_DETAIL_TERMS = [
    "Clique para detalhar",
    "Detalhar",
]
# ============================

def safe(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:120] if len(s) > 120 else s

def normalize(s: str) -> str:
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

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

def get_click_targets(page, terms):
    """Retorna pontos clicáveis que contenham os termos informados."""
    if not terms:
        return []

    script = r"""
    (terms) => {
      const normalize = (txt) =>
        (txt || "")
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/\s+/g, " ")
          .trim()
          .toLowerCase();

      const wanted = (terms || []).map(normalize).filter(Boolean);
      if (!wanted.length) return [];

      const out = [];
      const seen = new Set();

      for (const el of Array.from(document.querySelectorAll("body *"))) {
        const raw = (el.innerText || el.textContent || "").trim();
        if (!raw) continue;

        const norm = normalize(raw);
        if (!wanted.some((w) => norm.includes(w))) continue;

        let clickable = el.closest(
          'button,a,[role="button"],[onclick],[tabindex],.qv-object,.qv-object-wrapper,.cell'
        );
        if (!clickable) clickable = el;

        const rect = clickable.getBoundingClientRect();
        if (rect.width < 60 || rect.height < 12) continue;
        if (
          rect.right < 0 ||
          rect.bottom < 0 ||
          rect.left > window.innerWidth ||
          rect.top > window.innerHeight
        ) continue;

        const style = window.getComputedStyle(clickable);
        if (
          style.display === "none" ||
          style.visibility === "hidden" ||
          Number(style.opacity || 1) === 0
        ) continue;

        const key = `${Math.round((rect.left + rect.width / 2) / 20)}:${Math.round((rect.top + rect.height / 2) / 20)}`;
        if (seen.has(key)) continue;
        seen.add(key);

        out.push({
          key,
          text: raw.slice(0, 90),
          x: Math.round(rect.left + rect.width / 2),
          y: Math.round(rect.top + rect.height / 2),
          top: rect.top,
          left: rect.left
        });
      }

      out.sort((a, b) => (a.top - b.top) || (a.left - b.left));
      return out;
    }
    """
    try:
        return page.evaluate(script, terms) or []
    except Exception as e:
        print(f"[AVISO] Falha ao mapear detalhes: {e}")
        return []

def back_to_om(page, om: str) -> bool:
    if click_if_exists(page, "Voltar"):
        wait_qlik(page, extra_ms=5000)
        return True

    # fallback por texto parcial/caixa diferente
    loc = page.get_by_text(re.compile(r"voltar", re.IGNORECASE))
    if loc.count() > 0:
        loc.first.click()
        wait_qlik(page, extra_ms=5000)
        return True

    # último recurso: clicar novamente na própria OM no menu
    if click_menu_item(page, om):
        wait_qlik(page, extra_ms=6000)
        return True

    return False

def capture_detail_pages(page, om: str, shots, idx: int) -> int:
    terms = []
    terms.extend(DETAIL_TERMS_BY_OM.get(om, []))
    terms.extend(GENERIC_DETAIL_TERMS)

    # dedup mantendo ordem
    seen = set()
    unique_terms = []
    for t in terms:
        nt = normalize(t)
        if not nt or nt in seen:
            continue
        seen.add(nt)
        unique_terms.append(t)

    targets = get_click_targets(page, unique_terms)
    if not targets:
        return idx

    print(f"[INFO] {om}: {len(targets)} detalhe(s) mapeado(s).")

    for n, t in enumerate(targets, start=1):
        try:
            page.mouse.click(t["x"], t["y"])
        except Exception:
            print(f"[AVISO] Não consegui clicar no detalhe #{n} em {om}.")
            continue

        wait_qlik(page, extra_ms=7000)

        detail_name = safe(t.get("text", "DETALHAR")) or "DETALHAR"
        png = os.path.join(
            TMP_DIR,
            f"page_{idx:03d}_{safe(om)}_{safe(detail_name)}_{n:02d}.png",
        )
        screenshot_page(page, png)
        shots.append(png)
        print(f"[OK] Capturada: {om} -> {detail_name}")
        idx += 1

        if not back_to_om(page, om):
            print(f"[AVISO] Não achei caminho de volta em {om}; interrompendo detalhes dessa OM.")
            break

    return idx

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

            idx = capture_detail_pages(page, om, shots, idx)

        context.close()
        browser.close()

    build_pdf(shots, out_pdf)
    print(f"\nOK - PDF no mesmo estilo gerado em: {out_pdf}")
    print(f"Imagens temporárias em: {TMP_DIR}")

if __name__ == "__main__":
    main()
