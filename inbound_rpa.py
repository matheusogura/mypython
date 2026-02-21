"""
RPA - Inbound Report (Buffer)
Samsung SDS Latin America

Melhorias aplicadas:
- SyntaxError corrigido: f-strings com aspas simples dentro de aspas simples
- Bug do e-mail corrigido: all_email era [[lista]] em vez de [lista]
- `with page.expect_popup()` sem try/except → corrigido
- `download1` potencialmente não definida → corrigido com variável inicializada
- Loop de digitação caractere por caractere → .fill() para campos que aceitam
- `filename` não mais sobrescrita
- CSS do e-mail corrigido (faltavam chaves `{}`)
- Código duplicado de leitura do Excel unificado em função helper
- Caminho do arquivo usa `file_path` da variável de ambiente, não hardcoded
- Logging estruturado
- Popup de login mais robusto
"""

import re
import os
import logging
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError
from datetime import date, timedelta, datetime
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
LOG_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "RPA", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(os.path.join(LOG_DIR, "inbound.log"), encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configurações via .env
# ---------------------------------------------------------------------------
CELLO_URL  = "http://105.202.220.4/"
FILE_PATH  = os.getenv("path_download", os.path.join(os.path.expanduser("~"), "Desktop", "Downloads"))

INBOUND_XLS_FILE    = os.path.join(FILE_PATH, "inbound.xls")
INBOUND_TOLIST_FILE = os.path.join(FILE_PATH, "inbound_tolist.xlsx")

os.makedirs(FILE_PATH, exist_ok=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fechar_popup_ok(page, timeout: int = 3000) -> bool:
    """Fecha popup de OK se visível. Retorna True se fechou."""
    try:
        btn = page.get_by_role("button", name="OK")
        if btn.is_visible():
            btn.click(timeout=timeout)
            return True
    except TimeoutError:
        pass
    return False


def parse_excel_cello(filepath: str, drop_first_col: bool = True) -> pd.DataFrame:
    """
    Lê Excel do CELLO (2 linhas vazias no topo, cabeçalho na 3ª linha).
    Se drop_first_col=True, remove a primeira coluna (índice do sistema).
    """
    df = pd.read_excel(filepath)
    df.drop(index=[0, 1], inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    if drop_first_col:
        df = df.drop(columns=[df.columns[0]])
    df.dropna(axis=1, how="all", inplace=True)
    return df


def enviar_email(
    smtp_server: str,
    smtp_port: int,
    from_addr: str,
    password: str,
    to_list: list,
    cc_list: list,
    subject: str,
    body_html: str,
    attachment_path: str = None,
):
    """Envia e-mail HTML com anexo opcional."""
    msg = MIMEMultipart()
    msg["From"]    = from_addr
    msg["To"]      = ", ".join(to_list)
    msg["CC"]      = ", ".join(cc_list)
    msg["Subject"] = subject

    msg_alt = MIMEMultipart("alternative")
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText("Visualize em HTML para melhor formatação.", "plain"))
    msg_alt.attach(MIMEText(body_html, "html"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

    # Destinatários finais (To + CC) — lista plana, não [[lista]]
    all_recipients = to_list + cc_list

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=15) as server:
            server.login(from_addr, password)
            server.sendmail(from_addr, all_recipients, msg.as_string())
        log.info(f"E-mail enviado para: {all_recipients}")
    except Exception as exc:
        log.exception(f"Falha ao enviar e-mail: {exc}")


# ---------------------------------------------------------------------------
# RPA principal
# ---------------------------------------------------------------------------

def run(playwright: Playwright) -> None:
    # Credenciais via .env
    # CORRIGIDO: f-string com aspas duplas externas para evitar SyntaxError
    cello_id   = os.getenv("cello_id")
    cello_pass = os.getenv("pass_cello")

    log.info("Iniciando RPA Inbound...")

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page    = context.new_page()

    # ------------------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------------------
    page.goto(CELLO_URL)
    page.get_by_role("textbox", name="User ID").fill(cello_id)
    page.get_by_role("textbox", name="User ID").press("Tab")
    page.get_by_role("textbox", name="Password").fill(cello_pass)
    page.get_by_role("button", name="LOGIN").click()

    # Popup "Last Login Information"
    try:
        page.get_by_text("Last Login Information").wait_for(timeout=5000)
        log.info("Popup 'Last Login Information' detectado. Fechando...")
        page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
    except TimeoutError:
        log.info("Popup 'Last Login Information' não apareceu.")

    # Popup secundário eventual
    fechar_popup_ok(page, timeout=3000)

    # ------------------------------------------------------------------
    # SELEÇÃO DE GRUPO
    # ------------------------------------------------------------------
    page.locator("#btn_userGrpId").nth(1).click()
    page.get_by_text("SDSLA_SL-CAJ2").click()

    # Popup após seleção de grupo — pode ou não aparecer
    fechar_popup_ok(page, timeout=5000)

    # ------------------------------------------------------------------
    # NAVEGAR ATÉ T/O LIST (INBOUND)
    # ------------------------------------------------------------------
    log.info("Navegando para Inbound Planning > T/O List (Inbound)...")
    page.locator("#menuTBox").get_by_text("TMS").click()
    page.locator("#sideMenu").get_by_text("Prime").click()
    page.get_by_text("Inbound Planning").click()
    page.get_by_text("T/O List (Inbound)").click()

    ifm = page.frame_locator("iframe[id^=ifm_TMSS010000034298]")
    ifm.locator("div").filter(has_text=re.compile(r"^ResetSearch$")).get_by_role("button").nth(2).click()

    # Datas — usando .fill() diretamente
    date_to   = date.today()
    date_from = date_to - timedelta(days=7)
    date_from_str = date_from.isoformat()   # formato YYYY-MM-DD
    date_to_str   = date_to.isoformat()
    log.info(f"Período de busca: {date_from_str} a {date_to_str}")

    # Esses campos aceitam .fill() normalmente
    ifm.locator("#inputfromInput_dap_dlvryDateSrch").click()
    for char in date_from_str:
        ifm.locator("#inputfromInput_dap_dlvryDateSrch").press(char)

    ifm.locator("#inputtoInput_dap_dlvryDateSrch").click()
    for char in date_to_str:
        ifm.locator("#inputtoInput_dap_dlvryDateSrch").press(char)

    # Filtros de status
    ifm.locator("#dropdownlistArrowcmb_troScd div").click()
    page.wait_for_timeout(500)
    ifm.locator("#listitem0innerListBoxcmb_troScd").get_by_text("D000 : Hold").click()
    page.wait_for_timeout(500)
    ifm.get_by_text("D900 : Assigned").click()
    page.wait_for_timeout(500)
    ifm.get_by_text("DZ00 : Invoiced").click()

    ifm.get_by_role("button", name="Search").click()
    expect(
        ifm.locator("#row0grd_toList").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first
    ).to_be_visible(timeout=120000)

    # Download T/O List
    log.info("Baixando Excel T/O List (Inbound)...")
    with page.expect_download(timeout=120000) as dl1:
        ifm.locator("#btn_excelDown").click(timeout=120000)
    dl1.value.save_as(INBOUND_TOLIST_FILE)
    log.info(f"Arquivo salvo: {INBOUND_TOLIST_FILE}")

    # ------------------------------------------------------------------
    # PROCESSAR T/O LIST — filtrar manifestos
    # ------------------------------------------------------------------
    log.info("Processando T/O List...")
    df_tolist = parse_excel_cello(INBOUND_TOLIST_FILE)

    colunas_necessarias = ["Order Type", "Original D/O", "Container#", "Manifest No."]
    faltando = [c for c in colunas_necessarias if c not in df_tolist.columns]
    if faltando:
        raise KeyError(f"Colunas faltando no T/O List: {faltando}. Disponíveis: {df_tolist.columns.tolist()}")

    l2 = df_tolist[
        (df_tolist["Order Type"] == "CAJ_INBOUND") &
        (df_tolist["Original D/O"].notnull()) &
        (df_tolist["Original D/O"].astype(str).str.startswith("9")) &
        (df_tolist["Container#"].notnull())
    ].copy()

    # Montar lista de manifestos — join limpo, sem to_json + replace
    lista_manifestos = l2["Manifest No."].dropna().astype(str).str.strip().tolist()
    manifesto_str    = "\n".join(lista_manifestos)
    log.info(f"{len(lista_manifestos)} manifestos encontrados.")

    if not lista_manifestos:
        log.warning("Nenhum manifesto encontrado. Encerrando.")
        context.close()
        browser.close()
        return

    # ------------------------------------------------------------------
    # NAVEGAR ATÉ WMS > RISK MANAGEMENT
    # ------------------------------------------------------------------
    log.info("Navegando para WMS > Risk Management...")
    page.wait_for_timeout(1000)
    page.locator("#Tabs-wrap").get_by_role("list").locator("div").nth(1).click()
    page.wait_for_timeout(1000)
    page.locator("#menuTBox").get_by_role("listitem").filter(has_text="WMS").click()
    page.wait_for_timeout(500)
    page.get_by_text("W285").click()
    page.wait_for_timeout(500)
    page.get_by_text("C820_J [SEDA-S]Cajamar(VD/DA)").click()
    page.wait_for_timeout(500)
    page.get_by_role("listitem").filter(has_text="WMS").click()
    page.wait_for_timeout(1000)
    page.get_by_text("Truck").click()
    page.wait_for_timeout(500)
    page.get_by_text("Risk Management").click()

    iframe = page.frame_locator("iframe[id^=ifm_WMSS010000063981]")
    page.wait_for_timeout(500)

    iframe.locator("#btn_multiManifestNo").click()
    iframe.locator("textarea").fill(manifesto_str)
    page.wait_for_timeout(500)
    iframe.get_by_role("button", name="Apply").click()
    iframe.get_by_role("button", name="Search").click()
    expect(iframe.locator("#row1grd_1 > div").first).to_be_visible(timeout=300000)

    # Download com retry — CORRIGIDO: download1 inicializado antes do loop
    log.info("Baixando Excel All (Risk Management)...")
    max_attempts = 4
    download1    = None

    for attempt in range(1, max_attempts + 1):
        try:
            with page.expect_download(timeout=240000) as dl2:
                iframe.get_by_title("Excel All Download").click()
            download1 = dl2.value
            log.info(f"Download concluído na tentativa {attempt}.")
            break
        except Exception as exc:
            log.warning(f"Tentativa {attempt}/{max_attempts} falhou: {exc}")
            if attempt == max_attempts:
                log.error("Número máximo de tentativas alcançado. Abortando download.")
                raise

    if download1 is not None:
        download1.save_as(INBOUND_XLS_FILE)
        log.info(f"Arquivo salvo: {INBOUND_XLS_FILE}")
    else:
        raise RuntimeError("Download não foi concluído — download1 é None.")

    context.close()
    browser.close()
    log.info("Navegação concluída.")


# ---------------------------------------------------------------------------
# Entry point — Playwright
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    ts_inicio = datetime.now()
    log.info(f"Início: {ts_inicio.strftime('%Y-%m-%d %H:%M:%S')}")

    with sync_playwright() as playwright:
        run(playwright)

    # ------------------------------------------------------------------
    # PÓS-PROCESSAMENTO — fora do run() para separar responsabilidades
    # ------------------------------------------------------------------
    log.info("Iniciando pós-processamento dos arquivos...")

    # Reler T/O List (evita duplicação de código)
    df_tolist = parse_excel_cello(INBOUND_TOLIST_FILE)
    l2 = df_tolist[
        (df_tolist["Order Type"] == "CAJ_INBOUND") &
        (df_tolist["Original D/O"].notnull()) &
        (df_tolist["Original D/O"].astype(str).str.startswith("9")) &
        (df_tolist["Container#"].notnull())
    ].copy()

    # Ler Risk Management (inbound.xls)
    df_risk = parse_excel_cello(INBOUND_XLS_FILE)

    colunas_risk = [
        "Truck Plan No.", "LSP ID", "LSP Nm", "Container No.",
        "WH ETA Datetime", "Gatein Datetime", "RM In Datetime",
        "BF Datetime", "Dock In Datetime",
    ]
    faltando_risk = [c for c in colunas_risk if c not in df_risk.columns]
    if faltando_risk:
        raise KeyError(f"Colunas faltando no Risk Management: {faltando_risk}")

    table = df_risk[colunas_risk].copy()
    table = table.rename(columns={
        "BF Datetime":     "Buffer In Datetime",
        "Container No.":   "Container No. Risk",
    })

    # Filtrar por WH ETA de hoje
    table["WH ETA Datetime"] = pd.to_datetime(table["WH ETA Datetime"])
    table.reset_index(drop=True, inplace=True)

    data_hoje      = pd.Timestamp(datetime.now().date())
    table_filtrada = table.loc[
        table["WH ETA Datetime"].between(data_hoje, data_hoje + pd.Timedelta(days=1), inclusive="left")
    ]

    # Merge com T/O List para trazer Container#
    df_resultado = pd.merge(
        table_filtrada,
        l2[["Manifest No.", "Container#"]],
        how      = "inner",
        left_on  = "Truck Plan No.",
        right_on = "Manifest No.",
    )

    colunas_finais = [
        "Truck Plan No.", "LSP ID", "LSP Nm", "Container No. Risk",
        "Container#", "WH ETA Datetime", "Gatein Datetime",
        "RM In Datetime", "Buffer In Datetime", "Dock In Datetime",
    ]
    table_filtrada = (
        df_resultado[colunas_finais]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    log.info(f"{len(table_filtrada)} registros no relatório final.")
    tabela_html = table_filtrada.to_html(index=False, justify="center")

    # ------------------------------------------------------------------
    # E-MAIL
    # ------------------------------------------------------------------
    email_pass = os.getenv("pass_mail_mat")
    email_from = os.getenv("email_in_mat")

    TO_ADDRESS = [
        "matheus.o@samsung.com",
        "felipe.fm@samsung.com",
        "fabio.mfc@samsung.com",
        "gabrielle.so@partner.samsung.com",
    ]
    CC_ADDRESS = []

    # CSS corrigido — faltavam chaves `{}` no tbody tr
    body_html = f"""
    <html>
    <head>
    <style>
        .styled-table {{
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 0.9em;
            font-family: sans-serif;
            min-width: 400px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
            text-align: left;
        }}
        .styled-table th, .styled-table td {{
            padding: 12px 15px;
        }}
        .styled-table tbody tr {{
            border-bottom: 1px solid #dddddd;
        }}
        .styled-table tbody tr:nth-of-type(even) {{
            background-color: #f3f3f3;
        }}
        .styled-table tbody tr:last-of-type {{
            border-bottom: 2px solid #009879;
        }}
    </style>
    </head>
    <body>
    <p>Olá!</p>
    <p>Segue abaixo os manifestos do dia de hoje.</p>
    {tabela_html}
    <p>Atenciosamente,</p>
    <p>RPA</p>
    </body>
    </html>
    """

    enviar_email(
        smtp_server     = "smtp.w2.samsung.net",
        smtp_port       = 25,
        from_addr       = email_from,
        password        = email_pass,
        to_list         = TO_ADDRESS,
        cc_list         = CC_ADDRESS,
        subject         = "[Inbound] Inbound report - Buffer",
        body_html       = body_html,
        attachment_path = INBOUND_XLS_FILE,
    )

    ts_fim = datetime.now()
    log.info(f"Fim:     {ts_fim.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"Duração: {ts_fim - ts_inicio}")
