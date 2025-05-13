import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ────────────── Funções utilitárias da automação ──────────────


def esperar(driver, by, seletor, desc="", timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, seletor))
        )
    except TimeoutException:
        print(f"❌ Timeout ao esperar {desc or seletor}")
        return None


def clicar_js(driver, elemento, desc="elemento"):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", elemento)
    driver.execute_script("arguments[0].click();", elemento)
    print(f"🖱️  Cliquei em {desc}")


def nova_janela(driver, handles_antes, timeout=10):
    WebDriverWait(driver, timeout).until(
        lambda d: len(d.window_handles) > len(handles_antes)
    )
    novas = [h for h in driver.window_handles if h not in handles_antes]
    return novas[0] if novas else None


def clicar_emitir_nf(driver):
    xpath = (
        "//span[translate(normalize-space(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='emitir nota fiscal']/ancestor::button"
        "| //a[translate(normalize-space(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='emitir nota fiscal']"
    )
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.8)
    try:
        btn = driver.find_element(By.XPATH, xpath)
        clicar_js(driver, btn, "'Emitir Nota Fiscal'")
        return True
    except:
        pass
    altura_total = driver.execute_script("return document.body.scrollHeight;")
    y_atual = altura_total
    while y_atual > 0:
        y_atual -= 600
        driver.execute_script("window.scrollTo(0, arguments[0]);", y_atual)
        time.sleep(0.4)
        try:
            btn = driver.find_element(By.XPATH, xpath)
            clicar_js(driver, btn, "'Emitir Nota Fiscal'")
            return True
        except:
            continue
    return False


def escolher_todo_periodo(driver):
    print("🔄 Selecionando 'Todo o período'…")
    botao = esperar(
        driver, By.XPATH, '//button[contains(@class,"ds-loader-button__button")]', "botão de período")
    if not botao:
        return
    clicar_js(driver, botao, "botão de período")
    time.sleep(1.2)
    for item in driver.find_elements(By.CSS_SELECTOR, "div.ds-dropdown-item"):
        if "todo o período" in item.text.lower():
            clicar_js(driver, item, "'Todo o período'")
            time.sleep(2.5)
            return


def fechar_modal_antecipar(driver):
    try:
        modal = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH,
                                              "//div[contains(@class,'modal') or contains(@class,'ca-modal')][.//h2[contains(.,'Antecipar emissão')]]"
                                              ))
        )
        btn_cancelar = modal.find_element(
            By.CSS_SELECTOR, "button[data-cancel-button]")
        clicar_js(driver, btn_cancelar, "botão 'Cancelar' do modal")
        time.sleep(1.2)
        return True
    except TimeoutException:
        return False


def garantir_pagina_vendas(driver):
    url_desejada = "https://app.contaazul.com/#/ca/vendas/vendas-e-orcamentos"
    if not driver.current_url.startswith(url_desejada):
        print("↩️  Redirecionando para a lista de Vendas e Orçamentos…")
        driver.get(url_desejada)
        esperar(driver, By.CSS_SELECTOR, "table.ds-table", "tabela de vendas")
        time.sleep(2)

# ────────────── Automação principal ──────────────


def iniciar_automacao(path_planilha):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(
        ChromeDriverManager().install()), options=options)

    driver.get("https://pro.contaazul.com/")
    time.sleep(3)

    email = os.getenv(
        "CA_EMAIL", "XXXXXXXXX")
    senha = os.getenv("CA_PASSWORD", "XXXXXXXXX")

    esperar(driver, By.CSS_SELECTOR, 'input[type="email"]').send_keys(
        email, Keys.ENTER)
    time.sleep(3)
    esperar(driver, By.CSS_SELECTOR, 'input[type="password"]').send_keys(
        senha, Keys.ENTER)
    time.sleep(3)

    menu_vendas = esperar(
        driver, By.XPATH, '//div[.//span[text()="Vendas"] and contains(@class,"ds-row")]', "menu Vendas")
    clicar_js(driver, menu_vendas, "menu 'Vendas'")
    time.sleep(3)
    esperar(driver, By.ID, "SALES_CONTROL_BUDGETS_SALES").click()
    time.sleep(5)

    escolher_todo_periodo(driver)

    # carrega planilha e extrai lista de vendas
    try:
        wb = openpyxl.load_workbook(path_planilha)
        sheet = wb.active
        vendas = []
        for row in sheet.iter_rows(min_row=2, min_col=8, max_col=12):
            col_h = str(row[0].value).strip() if row[0].value else ""
            col_l = row[4].value if row[4].value else 0
            if col_h.lower().startswith("venda") and "/" not in col_h and ":" not in col_h:
                partes = col_h.split()
                if len(partes) == 2 and partes[1].isdigit() and float(col_l) > 0:
                    vendas.append(partes[1])
    except Exception as e:
        print("❌ Erro ao ler planilha:", e)
        driver.quit()
        return

    for venda in vendas:
        print(f"\n🔎 Buscando venda {venda}…")
        campo = esperar(driver, By.CSS_SELECTOR,
                        'input[placeholder="Pesquisar"]')
        campo.clear()
        campo.send_keys(venda, Keys.ENTER)
        time.sleep(20)

        esperar(driver, By.CSS_SELECTOR, "table.ds-table")
        time.sleep(2)

        achou = False
        for linha in driver.find_elements(By.CSS_SELECTOR, "table.ds-table tbody tr"):
            if linha.find_element(By.XPATH, "./td[3]").text.strip() != venda:
                continue

            achou = True
            # se já tiver SN, pula
            if "SN -" in linha.find_element(By.XPATH, "./td[8]").text:
                print(f"🔁 {venda} já possui SN – pulando.")
                break

            # ── AQUI substituímos a lógica: clicamos direto no link "Emitir NFS-e" ──
            try:
                print("🔎 Procurando link 'Emitir NFS-e' na linha…")
                handles_pre = driver.window_handles.copy()
                link_emitir = linha.find_element(
                    By.XPATH, ".//a[normalize-space(text())='Emitir NFS-e']")
                clicar_js(driver, link_emitir, "'Emitir NFS-e'")
                time.sleep(2)
            except Exception as e:
                print(f"❌ Erro ao tentar clicar em 'Emitir NFS-e': {e}")
                break

            # fecha o modal de antecipação, se aparecer
            if fechar_modal_antecipar(driver):
                print("ℹ️  Modal 'Antecipar emissão' cancelado.")

            aba_principal = driver.current_window_handle
            nova_aba = None
            try:
                nova_aba = nova_janela(driver, handles_pre, 10)
            except TimeoutException:
                pass

            if nova_aba:
                driver.switch_to.window(nova_aba)
                print("🔀 Nova aba aberta para emissão.")
            else:
                time.sleep(8)

            print("🔎 Procurando botão 'Emitir Nota Fiscal'…")
            if clicar_emitir_nf(driver):
                print(f"✅ Nota fiscal emitida para {venda}")
            else:
                print(f"❌ Não localizei 'Emitir Nota Fiscal' para {venda}")

            # fecha aba de emissão e volta
            if nova_aba:
                driver.close()
                driver.switch_to.window(aba_principal)
                print("⬅️ Fechei a aba de emissão.")
            else:
                driver.back()
                print("⬅️ Voltei para a listagem.")

            time.sleep(4)
            garantir_pagina_vendas(driver)
            escolher_todo_periodo(driver)
            time.sleep(1.5)
            break

        if not achou:
            print(f"⚠️ Venda {venda} não encontrada.")

    print("\n🏁 Processo concluído.")
    input("Pressione Enter para sair…")
    driver.quit()

# ────────────── Interface gráfica com Tkinter ──────────────


caminho_excel = ""


def selecionar_arquivo():
    global caminho_excel
    caminho_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )
    if caminho_excel:
        lbl_arquivo.config(
            text=f"Selecionado: {os.path.basename(caminho_excel)}")
    else:
        lbl_arquivo.config(text="Nenhum arquivo selecionado")


def executar_script():
    if not caminho_excel:
        messagebox.showwarning(
            "Arquivo necessário", "Por favor, selecione um arquivo Excel primeiro.")
        return
    threading.Thread(target=iniciar_automacao, args=(
        caminho_excel,), daemon=True).start()


janela = tk.Tk()
janela.title("Automação ContaAzul")
janela.geometry("400x200")

tk.Label(janela, text="1. Selecione o arquivo Excel (.xlsx):").pack(pady=(15, 5))
tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo).pack()
lbl_arquivo = tk.Label(janela, text="Nenhum arquivo selecionado", fg="gray")
lbl_arquivo.pack(pady=5)

tk.Label(janela, text="2. Clique para iniciar a automação:").pack(pady=(15, 5))
tk.Button(janela, text="Iniciar Automação",
          command=executar_script, bg="#00b894", fg="white").pack()

janela.mainloop()
