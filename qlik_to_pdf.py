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

def click_stage_action_by_text(page, text: str, nth: int = 0) -> bool:
    """
    Fallback robusto para clicar cards/ações no stage do Qlik, com match
    case/acento-insensível no texto renderizado do card.
    """
    target = normalize(text)
    if not target:
        return False

    try:
        return bool(page.evaluate(
            """({ target, nth }) => {
                const norm = (s) => (s || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .toLowerCase()
                    .replace(/\\s+/g, " ")
                    .trim();

                const isVisible = (el) => {
                    if (!el) return false;
                    const r = el.getBoundingClientRect();
                    if (!r || r.width < 1 || r.height < 1) return false;
                    const st = window.getComputedStyle(el);
                    return st.display !== "none" && st.visibility !== "hidden";
                };

                const root = document.querySelector("#qv-stage-container") || document;
                const cards = Array.from(root.querySelectorAll("article.qv-object"));
                const matches = cards.filter((card) => {
                    if (!isVisible(card)) return false;
                    const txt = norm(card.innerText || card.textContent || "");
                    return txt.includes(target);
                });

                const picked = matches[nth];
                if (!picked) return false;

                picked.scrollIntoView({ block: "center", inline: "center" });
                const btn = picked.querySelector("button");
                if (btn && !btn.disabled) {
                    btn.click();
                    return true;
                }
                picked.click();
                return true;
            }""",
            {"target": target, "nth": nth},
        ))
    except:
        return False

def wait_qlik(page, extra_ms=0):
    """
    Espera o Qlik estabilizar SEM depender de networkidle (que costuma nunca ficar idle).
    Critério: sumir loaders/spinners e dar um settle curto.
    """
    # 1) dá um micro-tempo para iniciar transição
    page.wait_for_timeout(250 + min(extra_ms, 800))

    # 2) espera loaders comuns do Qlik sumirem (ajuste seletores se necessário)
    loaders = [
        ".qv-loading",
        ".qv-loader",
        ".qv-spinner",
        ".lui-loading",
        "[class*='loading']",
        "[class*='spinner']",
    ]

    # tenta alguns passes rápidos (sem travar 12s)
    for _ in range(3):
        try:
            page.wait_for_function(
                """(sels) => sels.every(s => document.querySelectorAll(s).length === 0)""",
                arg=loaders,
                timeout=2500  # curto
            )
            break
        except PWTimeout:
            pass

    # 3) settle final pequeno (para garantir render antes do screenshot)
    page.wait_for_timeout(350 + min(extra_ms, 1200))

def click_menu_item(page, name: str) -> bool:
    rx_name = re.compile(rf"^\s*{re.escape(name)}\s*$", re.IGNORECASE)

    # espera um pouco o DOM do menu existir (sem depender de networkidle)
    try:
        page.wait_for_timeout(300)
        # se o seu menu é por role=button com name, isso ajuda o playwright a “ver” o elemento
        page.get_by_role("button", name=rx_name).first.wait_for(state="visible", timeout=2500)
    except:
        pass

    loc = page.get_by_role("button", name=rx_name)
    if loc.count() > 0:
        loc.first.click()
        return True

    if click_stage_action_by_text(page, name):
        return True

    stage_loc = page.locator("#qv-stage-container").locator(f"text={name}")
    if stage_loc.count() > 0:
        stage_loc.first.click()
        return True

    loc2 = page.locator(f"text={name}")
    if loc2.count() > 0:
        loc2.first.click()
        return True

    print(f"[AVISO] Menu '{name}' não encontrado.")
    return False

def click_text(page, text: str, nth: int = 0) -> bool:
    stage = page.locator("#qv-stage-container")

    # tenta clicar como footnote (muito comum no Qlik)
    foot = stage.locator("footer.qv-object-footnote", has_text=re.compile(re.escape(text), re.IGNORECASE))
    if foot.count() > nth:
        art = foot.nth(nth).locator("xpath=ancestor::article[contains(@class,'qv-object')]")
        if art.count() > 0:
            art.first.click()
        else:
            foot.nth(nth).click()
        return True

    # fallback: texto normal
    loc = stage.get_by_text(text, exact=False)
    if loc.count() > nth:
        loc.nth(nth).click()
        return True

    if click_stage_action_by_text(page, text, nth=nth):
        return True

    return False

def click_button(page, name: str, nth: int = 0) -> bool:
    loc = page.get_by_role("button", name=name)
    if loc.count() > nth:
        loc.nth(nth).click()
        return True
    return False

def click_by_bg_image(page, img_substring: str, nth: int = 0) -> bool:
    """
    Clica em botões do Qlik que não têm texto, mas têm background-image no style.
    Ex.: click1.png
    """
    raw = (img_substring or "").strip()
    if not raw:
        return False

    # Tentativa 1: seletor direto (mais rápido)
    for sel in (f'#qv-stage-container button[style*="{raw}"]', f'button[style*="{raw}"]'):
        loc = page.locator(sel)
        if loc.count() > nth:
            loc.nth(nth).click(force=True)
            return True

    # Tentativa 2: match case-insensitive + decode de URL + background computado
    token_base = raw.lower()
    base_name = os.path.basename(token_base)
    stem = base_name.rsplit(".", 1)[0]
    tokens = {token_base, base_name, stem}
    tokens.update(part for part in re.split(r"[_\-\s]+", stem) if len(part) >= 3)
    tokens = [t for t in tokens if t]

    try:
        return bool(page.evaluate(
            """({ tokens, nth }) => {
                const isVisible = (el) => {
                    if (!el) return false;
                    const r = el.getBoundingClientRect();
                    if (!r || r.width < 1 || r.height < 1) return false;
                    const st = window.getComputedStyle(el);
                    return st.display !== "none" && st.visibility !== "hidden";
                };

                const expand = (s) => {
                    const low = (s || "").toLowerCase();
                    try {
                        return low + " " + decodeURIComponent(low);
                    } catch {
                        return low;
                    }
                };

                const root = document.querySelector("#qv-stage-container") || document;
                const buttons = Array.from(root.querySelectorAll("article.qv-object button, button"));
                const matches = buttons.filter((btn) => {
                    if (!isVisible(btn)) return false;
                    const inlineStyle = expand(btn.getAttribute("style") || "");
                    const computedBg = expand(window.getComputedStyle(btn).backgroundImage || "");
                    const blob = inlineStyle + " " + computedBg;
                    return tokens.some((t) => blob.includes(t));
                });

                const picked = matches[nth];
                if (!picked) return false;
                picked.scrollIntoView({ block: "center", inline: "center" });
                picked.click();
                return true;
            }""",
            {"tokens": tokens, "nth": nth},
        ))
    except:
        return False

def back(page) -> bool:
    # Prioridade: voltar no stage, evitando o "Voltar uma etapa" da barra superior
    if click_stage_action_by_text(page, "Voltar"):
        wait_qlik(page, extra_ms=3500)
        return True

    stage = page.locator("#qv-stage-container")
    loc = stage.get_by_role("button", name=re.compile(r"^\s*voltar\s*$", re.IGNORECASE))
    if loc.count() > 0:
        loc.first.click()
        wait_qlik(page, extra_ms=3500)
        return True

    foot = stage.locator("footer.qv-object-footnote", has_text=re.compile(r"^\s*voltar\s*$", re.IGNORECASE))
    if foot.count() > 0:
        art = foot.first.locator("xpath=ancestor::article[contains(@class,'qv-object')]")
        if art.count() > 0:
            art.first.click()
        else:
            foot.first.click()
        wait_qlik(page, extra_ms=3500)
        return True

    return False

def back_to_om(page, om: str) -> bool:
    """
    Volta para a OM de forma robusta.
    1) tenta Voltar
    2) se não tiver, força clicando no menu da OM
    """
    if back(page):
        return True

    if click_menu_item(page, om):
        wait_qlik(page, extra_ms=5000)
        return True

    return False

def screenshot_page(page, out_png: str):
    """
    Captura a tela começando a partir do logo da página (sheet-title-logo-img).
    Recorta tudo que estiver acima do logo.
    """

    # espera garantir render
    page.wait_for_timeout(1500)

    logo = page.locator("div.sheet-title-logo-img")

    if logo.count() > 0:
        box = logo.first.bounding_box()

        if box:
            viewport = page.viewport_size

            clip = {
                "x": 0,
                "y": box["y"],  # começa no topo do logo
                "width": viewport["width"],
                "height": viewport["height"] - box["y"]
            }

            page.screenshot(path=out_png, clip=clip)
            return

    # fallback: se não achar o logo, captura normal
    page.screenshot(path=out_png, full_page=False)

def add_shot(page, shots, idx, label):
    png = os.path.join(TMP_DIR, f"page_{idx:03d}_{safe(label)}.png")
    screenshot_page(page, png)
    shots.append(png)
    print(f"[OK] Capturada: {label}")
    return idx + 1

def build_pdf(png_paths, out_pdf):
    if not png_paths:
        print("[ERRO] Nenhuma imagem foi gerada. Veja o PNG de debug e os avisos do log.")
        return

    imgs = [Image.open(p).convert("RGB") for p in png_paths]
    first, rest = imgs[0], imgs[1:]
    first.save(out_pdf, save_all=True, append_images=rest)

def click_card_like(page, contains_text: str, nth: int = 0) -> bool:
    """
    Qlik action button: o texto geralmente está no <footer class="qv-object-footnote">,
    e o botão clicável é um <button> sem texto (background-image).
    Então: clica no footer e/ou no article pai do objeto.
    """
    txt = (contains_text or "").strip()
    if not txt:
        return False

    stage = page.locator("#qv-stage-container")
    txt_rx = re.compile(re.escape(txt), re.IGNORECASE)

    # 1) Prioridade: footer (texto do card)
    foot = stage.locator("footer.qv-object-footnote", has_text=txt_rx)
    if foot.count() > nth:
        # clicar no article pai (área inteira clicável)
        art = foot.nth(nth).locator("xpath=ancestor::article[contains(@class,'qv-object')]")
        if art.count() > 0:
            art.first.click()
        else:
            foot.nth(nth).click()
        return True

    # 2) Fallback: qualquer elemento com o texto (bem amplo)
    anytxt = stage.locator(f"text={txt}")
    if anytxt.count() > nth:
        # tenta subir até o article do objeto e clicar
        art2 = anytxt.nth(nth).locator("xpath=ancestor::article[contains(@class,'qv-object')]")
        if art2.count() > 0:
            art2.first.click()
        else:
            anytxt.nth(nth).click()
        return True

    # 3) Fallback com normalização de acento/case
    if click_stage_action_by_text(page, txt, nth=nth):
        return True

    return False

def open_card(page, text_options=None, image_options=None) -> bool:
    text_options = text_options or []
    image_options = image_options or []

    for txt in text_options:
        if click_card_like(page, txt):
            return True

    for img in image_options:
        if click_by_bg_image(page, img, nth=0):
            return True

    return False

def main():
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    Path(TMP_DIR).mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_pdf = os.path.join(OUTPUT_DIR, f"e-GovEns_{ts}.pdf")

    # limpa tmp antigo
    for f in Path(TMP_DIR).glob("page_*.png"):
        try:
            f.unlink()
        except:
            pass

    shots = []
    idx = 1

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            channel="chrome"
        )
        context = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
            device_scale_factor=1
        )
        page = context.new_page()

        page.goto(QLIK_URL, wait_until="domcontentloaded", timeout=WAIT_MAX_MS)
        wait_qlik(page, extra_ms=4500)

        # CHECKPOINT: se nada funcionar, ao menos 1 captura para diagnóstico
        idx = add_shot(page, shots, idx, "00 - HOME (debug)")
        
        # =========================
        # ROTEIRO FIXO = PDF ANEXADO
        # =========================

        # --- AFA (11 páginas) ---
        if click_menu_item(page, "AFA"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "AFA")

            # Tabs topo na ordem do PDF
            for tab in ["Formação Aviadores", "Formação Intendência", "Formação Infantaria", "Aeronave / Simulador", "Aeronave/Simulador"]:
                # tentamos as 2 grafias do Simulador, mas só uma vai existir
                if "Aeronave" in tab:
                    if click_text(page, "Aeronave / Simulador") or click_text(page, "Aeronave/Simulador"):
                        wait_qlik(page, extra_ms=5500)
                        idx = add_shot(page, shots, idx, "AFA - Aeronave_Simulador")
                    continue

                if click_text(page, tab):
                    wait_qlik(page, extra_ms=5500)
                    idx = add_shot(page, shots, idx, f"AFA - {tab}")

            # Instrução T25 (principal)
            if click_text(page, "Instrução T25"):
                wait_qlik(page, extra_ms=6000)
                idx = add_shot(page, shots, idx, "AFA - Instrução T25")

                # Existem 2 botões "Detalhar" no T25 (CFOAV 2º e 4º) – PDF tem os 2
                # Clique no 1º Detalhar
                if click_button(page, "Detalhar", nth=0) or click_text(page, "Detalhar", nth=0):
                    wait_qlik(page, extra_ms=7000)
                    idx = add_shot(page, shots, idx, "AFA - T25 Detalhar CFOAV 2º Ano")
                    back(page)

                # Clique no 2º Detalhar
                if click_button(page, "Detalhar", nth=1) or click_text(page, "Detalhar", nth=1):
                    wait_qlik(page, extra_ms=7000)
                    idx = add_shot(page, shots, idx, "AFA - T25 Detalhar CFOAV 4º Ano")
                    back(page)

            # Instrução T27 (principal)
            if click_text(page, "Instrução T27"):
                wait_qlik(page, extra_ms=6000)
                idx = add_shot(page, shots, idx, "AFA - Instrução T27")

                # No T27 o PDF tem 1 Detalhar
                if click_button(page, "Detalhar", nth=0) or click_text(page, "Detalhar", nth=0):
                    wait_qlik(page, extra_ms=7000)
                    idx = add_shot(page, shots, idx, "AFA - T27 Detalhar CFOAV 4º Ano")
                    back(page)

            # Esforço Aéreo
            if click_text(page, "Esforço Aéreo"):
                wait_qlik(page, extra_ms=6000)
                idx = add_shot(page, shots, idx, "AFA - Esforço Aéreo")

        # --- EPCAR (2) ---
        if click_menu_item(page, "EPCAR"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "EPCAR")
            if click_card_like(page, "Cursos da EPCAR"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "EPCAR - Cursos")

        # --- EEAR (2) ---
        if click_menu_item(page, "EEAR"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "EEAR")
            if click_card_like(page, "Cursos da EEAR"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "EEAR - Cursos")

        # --- CIAAR (2) ---
        if click_menu_item(page, "CIAAR"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "CIAAR")
            if click_card_like(page, "Cursos do CIAAR") or click_card_like(page, "Cursos da CIAAR"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "CIAAR - Cursos")

        # --- IEAD (2) ---
        if click_menu_item(page, "IEAD"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "IEAD")
            if click_card_like(page, "Cursos do IEAD") or click_card_like(page, "Cursos da IEAD"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "IEAD - Cursos")

        # --- UNIFA (2) ---
        if click_menu_item(page, "UNIFA"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "UNIFA")
            if click_card_like(page, "Cursos da UNIFA") or click_card_like(page, "Cursos do UNIFA"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "UNIFA - Cursos")

        # --- ECEMAR (4) ---
        if click_menu_item(page, "ECEMAR"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "ECEMAR")

            # 1) Cursos da ECEMAR (PDF tem)
            if open_card(
                page,
                text_options=["Cursos da ECEMAR", "Cursos do ECEMAR", "Cursos ECEMAR"],
                image_options=["ecemar"],
            ):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "ECEMAR - Cursos")
                back_to_om(page, "ECEMAR")

            # 2) PLAMENS exterior (PDF tem 2 capturas seguidas dessa tela)
            if open_card(
                page,
                text_options=[
                    "PLAMENS exterior",
                    "PLAMENS Exterior",
                    "PLAMENS EXTERIOR",
                    "PLAMENS no exterior",
                ],
                image_options=["plamens"],
            ):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "ECEMAR - PLAMENS exterior (1)")

                # segunda captura
                try:
                    page.mouse.wheel(0, 700)
                    page.wait_for_timeout(1500)
                except:
                    pass
                idx = add_shot(page, shots, idx, "ECEMAR - PLAMENS exterior (2)")

                back_to_om(page, "ECEMAR")

        # --- EAOAR (2) ---
        if click_menu_item(page, "EAOAR"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "EAOAR")
            if click_card_like(page, "Cursos da EAOAR") or click_card_like(page, "Cursos do EAOAR"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "EAOAR - Cursos")

        # ==========================================================
        # BLOCO DIRENS TROCAD0 (conforme seu snippet, click1.png)
        # ==========================================================
        # --- DIRENS (2) ---
        if click_menu_item(page, "DIRENS"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "DIRENS")

            # subtela é um <button> com background-image click1.png (sem texto)
            if click_by_bg_image(page, "click1.png", nth=0) or click_card_like(page, "Dados dos Exames"):
                wait_qlik(page, extra_ms=6500)
                idx = add_shot(page, shots, idx, "DIRENS - Dados dos Exames")
                back_to_om(page, "DIRENS")
            else:
                print("[AVISO] DIRENS: não achei button com click1.png para abrir a subtela.")

               # ==========================================================
        # BLOCO ASSISTENCIAIS (4) — CBNB / CTRB / ECE por background-image
        # ==========================================================
        if click_menu_item(page, "ASSISTENCIAIS"):
            wait_qlik(page, extra_ms=6500)
            idx = add_shot(page, shots, idx, "ASSISTENCIAIS - Cards")

            assistenciais_cards = [
                (["CBNB"], ["cbnb_egovens_2.png", "cbnb"], "ASSISTENCIAIS - CBNB"),
                (["CTRB"], ["ctrb_egovens_2.png", "ctrb"], "ASSISTENCIAIS - CTRB"),
                (["ECE"], ["ece_egovens_2.png", "ece_egovens"], "ASSISTENCIAIS - ECE"),
            ]

            for txt_opts, img_opts, label in assistenciais_cards:
                if open_card(page, text_options=txt_opts, image_options=img_opts):
                    wait_qlik(page, extra_ms=6500)
                    idx = add_shot(page, shots, idx, label)
                    back_to_om(page, "ASSISTENCIAIS")
                else:
                    print(f"[AVISO] ASSISTENCIAIS: não achei card {label} (texto/imagem).")

        context.close()
        browser.close()

    build_pdf(shots, out_pdf)
    print(f"\nOK - PDF gerado em: {out_pdf}")
    print(f"Imagens temporárias em: {TMP_DIR}")
    print(f"Total de páginas: {len(shots)}")

if __name__ == "__main__":
    main()
