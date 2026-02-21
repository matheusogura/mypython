"""
RPA - Rate Shop Automation
Samsung SDS Latin America

Melhorias aplicadas:
- Credenciais via variáveis de ambiente
- Caminhos dinâmicos (não hardcoded)
- Logging estruturado
- Tratamento de exceções genéricas no loop principal
- Timeout explícito em todos os `expect`
- `dropna` corrigido para não remover colunas essenciais
- Leitura do DataFrame corrigida (índices e cabeçalho)
- `to_string()` substituído por `join` seguro
- Variáveis com nomes distintos para não sobrescrever data de início
- `time.sleep` substituído por `wait_for_timeout`
"""

import re
import os
import json
import logging
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError
from datetime import date, timedelta, datetime
import pandas as pd

# ---------------------------------------------------------------------------
# Configuração de logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            os.path.join(os.path.expanduser("~"), "Desktop", "RPA", "logs", "rateshop.log"),
            encoding="utf-8",
        ),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configurações — use variáveis de ambiente para credenciais!
# Ex: set CELLO_USER=matheus.o && set CELLO_PASS=SuaSenha
# ---------------------------------------------------------------------------
CELLO_URL  = os.getenv("CELLO_URL",  "http://105.202.220.4/cello/view/login.html?sso=N&idp=N")
CELLO_USER = os.getenv("CELLO_USER", "matheus.o")         # ⚠️ prefira sempre via env var
CELLO_PASS = os.getenv("CELLO_PASS", "Dezembro@@12")      # ⚠️ prefira sempre via env var

BASE_DIR        = os.path.join(os.path.expanduser("~"), "Desktop", "Downloads")
RESULTS_DIR     = os.path.join(os.path.expanduser("~"), "Desktop", "RPA", "resultados")
SHIPMENT_FILE   = os.path.join(BASE_DIR, "shipment.xlsx")
CTE_FILE        = os.path.join(BASE_DIR, "cte.xlsx")

# Garante que os diretórios existam
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)
os.makedirs(os.path.join(os.path.expanduser("~"), "Desktop", "RPA", "logs"), exist_ok=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_excel_with_header_offset(filepath: str) -> pd.DataFrame:
    """
    Lê um Excel do CELLO que tem 2 linhas vazias no topo antes do cabeçalho real.
    Retorna um DataFrame limpo.
    """
    df = pd.read_excel(filepath)
    df.drop(index=[0, 1], inplace=True)          # Remove linhas vazias iniciais
    df.columns = df.iloc[0]                       # Define cabeçalho
    df = df.iloc[1:].reset_index(drop=True)       # Remove a linha de cabeçalho dos dados

    # Remove apenas colunas completamente vazias (todas NaN), não qualquer NaN
    df.dropna(axis=1, how="all", inplace=True)
    return df


def fechar_popup_ok(page, timeout: int = 3000) -> bool:
    """Tenta fechar um popup de OK. Retorna True se fechou, False se não havia popup."""
    try:
        btn = page.get_by_role("button", name="OK")
        if btn.is_visible():
            btn.click(timeout=timeout)
            return True
    except TimeoutError:
        pass
    return False


# ---------------------------------------------------------------------------
# Função principal do RPA
# ---------------------------------------------------------------------------

def run(playwright: Playwright) -> None:
    log.info("Iniciando RPA Rate Shop...")
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page    = context.new_page()

    # ------------------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------------------
    log.info("Navegando para o CELLO...")
    page.goto(CELLO_URL)
    page.get_by_role("textbox", name="User ID").fill(CELLO_USER)
    page.get_by_role("textbox", name="User ID").press("Tab")
    page.get_by_role("textbox", name="Password").fill(CELLO_PASS)
    page.get_by_role("button", name="LOGIN").click()

    # Popup "Last Login Information"
    try:
        mensagem = page.get_by_text("Last Login Information")
        mensagem.wait_for(timeout=5000)
        log.info("Popup 'Last Login Information' detectado. Fechando...")
        page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
    except TimeoutError:
        log.info("Popup 'Last Login Information' não apareceu. Continuando...")

    # ------------------------------------------------------------------
    # SELEÇÃO DE GRUPO
    # ------------------------------------------------------------------
    page.locator("#btn_userGrpId").nth(1).click()
    page.get_by_text("SDSLA_SL-CAJ2").dblclick()
    page.wait_for_timeout(500)

    # Segundo popup (pode aparecer após seleção de grupo)
    fechar_popup_ok(page, timeout=3000)

    # ------------------------------------------------------------------
    # NAVEGAR ATÉ SHIPMENT
    # ------------------------------------------------------------------
    log.info("Navegando até TMS > Prime > Transport Planning > Shipment...")
    page.locator("#menuTBox").get_by_text("TMS").click(timeout=60000)
    page.locator("#sideMenu").get_by_text("Prime").click()
    page.get_by_text("Transport Planning").nth(1).click()
    page.get_by_role("paragraph").filter(has_text=re.compile(r"^Shipment$")).click()
    page.bring_to_front()

    # ------------------------------------------------------------------
    # FILTROS DE BUSCA — SHIPMENT
    # ------------------------------------------------------------------
    iframe = page.frame_locator("iframe[id^=ifm_TMSS010000026557]")
    iframe.locator("#dropdownlistArrowcmb_lspId div").click()
    page.wait_for_timeout(500)
    iframe.get_by_role("textbox", name="Looking for").fill("G8955")
    iframe.get_by_role("textbox", name="Looking for").press("Enter")
    page.wait_for_timeout(500)
    iframe.get_by_text("G8955 : TOTAL EXPRESS (E-").click()

    # Botão Reset
    iframe.get_by_text("ResetSearch").click()
    iframe.locator("div").filter(has_text=re.compile(r"^ResetSearch$")).get_by_role("button").nth(2).click()

    # Calcular intervalo de datas
    hoje       = date.today()
    dia_semana = hoje.weekday()
    if dia_semana == 0:  # Segunda-feira → vai para sexta anterior
        data_fim = hoje - timedelta(days=3)
    else:
        data_fim = hoje - timedelta(days=1)

    data_inicio_str = hoje.isoformat()
    data_fim_str    = data_fim.isoformat()
    log.info(f"Período de busca: {data_fim_str} a {data_inicio_str}")

    iframe.locator("#inputfromInput_dap_shmptCreatDate").fill(data_fim_str)
    iframe.locator("#inputtoInput_dap_shmptCreatDate").fill(data_inicio_str)
    iframe.locator("div").filter(has_text=re.compile(r"^Setting Reset Search$")).get_by_role("button").nth(3).click()
    iframe.get_by_role("button", name="Search").click()

    log.info("Aguardando resultados da busca de Shipment...")
    expect(
        iframe.locator("#row0grd_shmpt").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first
    ).to_be_visible(timeout=600000)

    # ------------------------------------------------------------------
    # DOWNLOAD — SHIPMENT EXCEL
    # ------------------------------------------------------------------
    log.info("Baixando Excel de Shipment...")
    with page.expect_download(timeout=600000) as download_info:
        iframe.get_by_title("Shipment Excel Download").click()
    download_info.value.save_as(SHIPMENT_FILE)
    log.info(f"Shipment salvo em: {SHIPMENT_FILE}")

    # ------------------------------------------------------------------
    # PROCESSAR SHIPMENT — filtrar Order_kgs entre 0.50 e 0.59
    # ------------------------------------------------------------------
    log.info("Processando arquivo de Shipment...")
    df_shipment = parse_excel_with_header_offset(SHIPMENT_FILE)

    if "Order_kgs" not in df_shipment.columns:
        raise KeyError(f"Coluna 'Order_kgs' não encontrada. Colunas disponíveis: {df_shipment.columns.tolist()}")

    df_filtrado  = df_shipment[(df_shipment["Order_kgs"] >= 0.50) & (df_shipment["Order_kgs"] < 0.59)]
    lista_shipments = df_filtrado["Shipment No."].dropna().astype(str).str.strip().tolist()
    sl1 = "\n".join(lista_shipments)   # Sem espaços extras (to_string() pode gerar padding)
    log.info(f"{len(lista_shipments)} shipments encontrados no filtro de peso.")

    if not lista_shipments:
        log.warning("Nenhum shipment encontrado no filtro. Encerrando.")
        context.close()
        browser.close()
        return

    # ------------------------------------------------------------------
    # NAVEGAR ATÉ CTE MONITORING ONLINE
    # ------------------------------------------------------------------
    log.info("Navegando para CTE Monitoring Online...")
    page.get_by_text("Interface", exact=True).click()
    page.get_by_text("CTE Monitoring Online").click()

    iframe_cte = page.frame_locator("iframe[id^=ifm_TMSS010000060958]")
    iframe_cte.locator("#btn_shmptNo").click(timeout=60000)
    iframe_cte.locator("textarea").fill(sl1)
    iframe_cte.get_by_role("button", name="Apply").click()
    iframe_cte.get_by_role("button", name="Search").click()

    log.info("Aguardando resultados do CTE...")
    expect(
        iframe_cte.locator("#row0grd_cteMntr").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first
    ).to_be_visible(timeout=300000)  # timeout explícito: 5 minutos

    # ------------------------------------------------------------------
    # DOWNLOAD — CTE EXCEL
    # ------------------------------------------------------------------
    log.info("Baixando Excel de CTE...")
    with page.expect_download(timeout=120000) as download_info:
        iframe_cte.get_by_title("Excel Download").click(timeout=60000)
    download_info.value.save_as(CTE_FILE)
    log.info(f"CTE salvo em: {CTE_FILE}")

    # ------------------------------------------------------------------
    # PROCESSAR CTE — filtrar Status == 'Ready'
    # ------------------------------------------------------------------
    log.info("Processando arquivo de CTE...")
    df_cte = parse_excel_with_header_offset(CTE_FILE)

    if "Status Name" not in df_cte.columns:
        raise KeyError(f"Coluna 'Status Name' não encontrada. Colunas disponíveis: {df_cte.columns.tolist()}")

    df_ready = df_cte[df_cte["Status Name"] == "Ready"]
    sl_ready = df_ready["Shipment No."].dropna().astype(str).str.strip().reset_index(drop=True)
    log.info(f"{len(sl_ready)} shipments com status 'Ready' encontrados.")

    # ------------------------------------------------------------------
    # RATE SHOP — processar cada shipment
    # ------------------------------------------------------------------
    log.info("Iniciando processo de Rate Shop...")
    page.locator("#Tabs-wrap").get_by_text("Shipment").click()
    iframe = page.frame_locator("iframe[id^=ifm_TMSS010000026557]")
    iframe.get_by_role("button", name="Reset").click()
    fechar_popup_ok(page)

    resultados = []

    for qty, shipment in enumerate(sl_ready):
        log.info(f"[{qty+1}/{len(sl_ready)}] Processando shipment: {shipment}")
        status = None

        try:
            iframe = page.frame_locator("iframe[id^=ifm_TMSS010000026557]")
            iframe.locator("#ipt_shmptNo").click()
            iframe.locator("#ipt_shmptNo").fill(shipment)
            iframe.get_by_role("button", name="Search").click()

            try:
                status = iframe.locator("xpath=//*[@id='row0grd_shmpt']/div[4]/div").inner_text(timeout=15000)
            except TimeoutError:
                log.warning(f"  Shipment {shipment} não encontrado (timeout na busca).")
                status = "shipment não encontrada"
                resultados.append({"qty": qty, "shipment": shipment, "status": status})
                continue

            if status != "Cons":
                log.info(f"  Status '{status}' — não é 'Cons', pulando Rate Shop.")
                resultados.append({"qty": qty, "shipment": shipment, "status": status})
                continue

            # Status é 'Cons' — executar Rate Shop
            iframe.locator("#columntablegrd_shmpt").get_by_role("checkbox").click(timeout=60000)
            iframe.get_by_text("Rate Shop").click(timeout=120000)

            try:
                popup = page.locator("#iframePopup").content_frame
                popup.get_by_text("TOTAL EXPRESS (E-store)").wait_for(timeout=90000)

                carrier_element = popup.locator('[role="gridcell"]:has-text("TOTAL EXPRESS (E-store)")').first
                row = carrier_element.locator('xpath=./ancestor::*[@role="row"]')

                # Tentar clicar na linha — estratégias em cascata
                clicked = False
                if row.count() > 0:
                    first_cell = row.locator('[role="gridcell"]').first
                    if first_cell.is_visible():
                        first_cell.click()
                        clicked = True
                        log.info("  Clicou na primeira célula da linha.")
                    else:
                        clickable = carrier_element.locator('[role="button"], button, [tabindex]')
                        if clickable.count() > 0:
                            clickable.click()
                            clicked = True
                            log.info("  Clicou no elemento clicável dentro do gridcell.")
                        else:
                            carrier_element.evaluate("el => el.click()")
                            clicked = True
                            log.info("  Clicou via JavaScript no gridcell.")

                if clicked:
                    popup.locator("#dropdownlistArrowcmb_reasonCd_tmsShmptRateShopP01 div").click()
                    page.wait_for_timeout(300)
                    popup.get_by_text("Weight Divergence").click()
                    popup.get_by_text("Save").nth(1).click()
                    fechar_popup_ok(page)
                    log.info(f"  Rate Shop aplicado com sucesso para {shipment}.")
                else:
                    log.warning(f"  Não foi possível selecionar a linha da carrier para {shipment}.")
                    popup.locator("xpath=/html/body/div[3]/a/img").click()
                    status = "Checkbox not found"

            except TimeoutError:
                log.warning(f"  Timeout ao buscar 'TOTAL EXPRESS' no popup para {shipment}.")
                try:
                    page.locator("#iframePopup").content_frame.locator("xpath=/html/body/div[3]/a/img").click()
                except TimeoutError:
                    fechar_popup_ok(page)
                status = "Carrier not found"

        except TimeoutError:
            log.error(f"  Timeout geral ao processar shipment {shipment}.")
            status = "shipment não encontrada"

        except Exception as exc:
            # Captura qualquer outra exceção para não crashar o loop inteiro
            log.exception(f"  Erro inesperado ao processar shipment {shipment}: {exc}")
            status = f"Erro inesperado: {exc}"

        resultados.append({"qty": qty, "shipment": shipment, "status": status})

    # ------------------------------------------------------------------
    # SALVAR RESULTADOS
    # ------------------------------------------------------------------
    df_resultado   = pd.DataFrame(resultados)
    data_atual_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    output_path    = os.path.join(RESULTS_DIR, f"rateshop_{data_atual_str}.txt")
    df_resultado.to_csv(output_path, sep=";", index=False)
    log.info(f"Resultados salvos em: {output_path}")

    context.close()
    browser.close()
    log.info("RPA finalizado com sucesso!")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    ts_inicio = datetime.now()
    log.info(f"Início da execução: {ts_inicio.strftime('%Y-%m-%d %H:%M:%S')}")

    with sync_playwright() as playwright:
        run(playwright)

    ts_fim = datetime.now()
    log.info(f"Fim da execução:    {ts_fim.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"Tempo total:        {ts_fim - ts_inicio}")
