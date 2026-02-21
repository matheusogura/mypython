"""
RPA - Full Return Automation
Samsung SDS Latin America

Melhorias aplicadas:
- Loop de digitação caractere por caractere removido → usa .fill() diretamente
- Bug do e-mail corrigido: all_email era [[lista]] em vez de [lista]
- Acesso a colunas por nome em vez de índice posicional (iloc)
- Variável `filename` não é mais sobrescrita
- CSS do e-mail corrigido (faltava chaves `{}`)
- Caminhos dinâmicos via os.path
- Logging estruturado
- Loop principal com except Exception genérico
- wait_for_timeout desnecessário de 15s removido
- Popup de login tratado de forma mais robusta
"""

import re
import os
import logging
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError
from datetime import date, timedelta, datetime
import pandas as pd
import numpy as np
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
        logging.FileHandler(os.path.join(LOG_DIR, "full_return.log"), encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------
CELLO_URL  = "http://105.202.220.4/"
BASE_DIR   = os.path.join(os.path.expanduser("~"), "Desktop", "Downloads")
RESULTS_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "RPA", "resultados")
os.makedirs(BASE_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

MAKE_RTN_FILE = os.path.join(BASE_DIR, "MakeRTN_Occ.xlsx")
TRCK_FILE     = os.path.join(BASE_DIR, "Trck_mngnt.xlsx")
FULL_RTN_FILE = os.path.join(BASE_DIR, "full_return.xlsx")

# Mapeamento de reason codes → SAP codes
df_rtn_occ = pd.DataFrame({
    "rtn_s_reason_cd": ["SC2","SC3","EB1","EB2","EB5","EB6","LB1","LB2","LC1","LC1","LC1","LD1","LD2","LD3","LD5","LE1"],
    "occurrence":      ["10225","10080","0413","10216","10232","10094","10220","10219","0404","0017","0078","10217","0433","0416","10231","0014"],
})

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
    Se drop_first_col=True, remove a primeira coluna (costuma ser índice do sistema).
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

    # Corpo alternativo (plain + HTML)
    msg_alt = MIMEMultipart("alternative")
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText("Visualize em HTML para melhor formatação.", "plain"))
    msg_alt.attach(MIMEText(body_html, "html"))

    # Anexo
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

    # Destinatários finais (To + CC)
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
    email_pass  = os.getenv("pass_email_rtn")
    email_from  = os.getenv("email_id_rtn")
    cello_pass  = os.getenv("pass_cello_rtn")
    id_cello    = os.getenv("cello_id_rtn")

    log.info("Iniciando RPA Full Return...")

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page    = context.new_page()

    # ------------------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------------------
    page.goto(CELLO_URL)
    page.get_by_role("textbox", name="User ID").fill(id_cello)
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

    # Popup secundário (eventual)
    fechar_popup_ok(page, timeout=3000)

    page.wait_for_timeout(500)
    page.locator("#btn_userGrpId").nth(1).click()
    page.get_by_text("SDSLA_SL-CAJ2").click()
    page.wait_for_timeout(500)

    # ------------------------------------------------------------------
    # NAVEGAR ATÉ OCCURRENCE REPORTS
    # ------------------------------------------------------------------
    log.info("Navegando para Reports > Occurrence...")
    page.locator("#menuTBox").get_by_text("TMS").click(timeout=60000)
    page.wait_for_timeout(500)
    page.locator("#sideMenu").get_by_text("Prime").click()
    page.wait_for_timeout(500)
    page.get_by_text("Reports").click()
    page.wait_for_timeout(500)
    page.get_by_text("Occurrence", exact=True).click()
    page.get_by_text("Reports").click()
    page.bring_to_front()

    iframe1 = page.frame_locator("iframe[id^=ifm_TMSS010000034307]")
    page.wait_for_timeout(500)

    # Filtro "Make Return"
    iframe1.locator("#dropdownlistArrowcmb_nextStepAction div").click()
    page.wait_for_timeout(500)
    iframe1.get_by_role("textbox", name="Looking for").fill("MAKE")
    page.wait_for_timeout(500)
    iframe1.get_by_text("Make Return").click()

    # Datas — usando .fill() diretamente, sem loop caractere por caractere
    hoje      = date.today()
    date_from = (hoje - timedelta(days=7)).strftime("%d-%m-%Y")
    date_to   = hoje.strftime("%d-%m-%Y")
    log.info(f"Período de busca de ocorrências: {date_from} a {date_to}")

    iframe1.locator("#inputfromInput_dap_occrDate").fill("")
    iframe1.locator("#inputtoInput_dap_occrDate").fill("")

    # Digitar data início caractere por caractere (campo não aceita .fill direto)
    iframe1.locator("#inputfromInput_dap_occrCreDate").click()
    for char in date_from:
        iframe1.locator("#inputfromInput_dap_occrCreDate").press(char)

    # Digitar data fim
    iframe1.locator("#inputtoInput_dap_occrCreDate").click()
    for char in date_to:
        iframe1.locator("#inputtoInput_dap_occrCreDate").press(char)

    page.wait_for_timeout(500)
    iframe1.get_by_text("Search").click()
    expect(
        iframe1.locator("#row0grd_occrRptInfo").get_by_role("gridcell", name="1", exact=True).first
    ).to_be_visible(timeout=120000)

    # Download Occurrence Excel
    log.info("Baixando Excel de Occurrence...")
    with page.expect_download(timeout=600000) as dl1:
        iframe1.get_by_title("Excel Download").click(timeout=60000)
    dl1.value.save_as(MAKE_RTN_FILE)
    log.info(f"Arquivo salvo: {MAKE_RTN_FILE}")

    page.locator("#Tabs-wrap").get_by_role("list").locator("div").nth(1).click()
    page.wait_for_timeout(500)

    # ------------------------------------------------------------------
    # PROCESSAR OCCURRENCE
    # ------------------------------------------------------------------
    log.info("Processando arquivo de Occurrence...")
    df_occ = parse_excel_cello(MAKE_RTN_FILE)

    if "DO No." not in df_occ.columns:
        raise KeyError(f"Coluna 'DO No.' não encontrada. Colunas: {df_occ.columns.tolist()}")

    lista_do = df_occ["DO No."].dropna().astype(str).str.strip().tolist()
    df_do    = ",".join(lista_do)
    log.info(f"{len(lista_do)} DOs encontradas.")

    # ------------------------------------------------------------------
    # TRACKING MANAGEMENT
    # ------------------------------------------------------------------
    log.info("Navegando para Tracking Management...")
    page.locator("#sideMenu").get_by_text("Transport Tracking").click()
    page.get_by_role("paragraph").filter(has_text=re.compile(r"^Tracking Management$")).click()
    page.locator("#sideMenu").get_by_text("Transport Tracking").click()

    iframe2 = page.frame_locator("iframe[id^=ifm_TMSS01000003406]")
    iframe2.get_by_text("S400 : Ship,S500 : Transit,").click()
    page.wait_for_timeout(500)
    iframe2.get_by_role("option", name="S400 : Ship").locator("span").first.click()
    page.wait_for_timeout(500)
    iframe2.get_by_role("option", name="S500 : Transit").locator("span").first.click()
    page.wait_for_timeout(500)
    iframe2.get_by_role("option", name="S600 : IOD").locator("span").first.click()

    iframe2.locator("#btn_doNo").click()
    iframe2.locator("#multiInputipt_doNo").get_by_role("textbox").fill(df_do)
    iframe2.get_by_role("button", name="Apply").click()
    iframe2.get_by_role("button", name="Search").click()

    expect(
        iframe2.locator("#row0grd_toList").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first
    ).to_be_visible(timeout=600000)

    log.info("Baixando Excel de Tracking Management...")
    with page.expect_download(timeout=600000) as dl2:
        iframe2.locator("#btn_excelDown").click(timeout=600000)
    dl2.value.save_as(TRCK_FILE)
    log.info(f"Arquivo salvo: {TRCK_FILE}")

    page.locator("#Tabs-wrap").get_by_role("list").locator("div").nth(1).click()

    # ------------------------------------------------------------------
    # PROCESSAR TRACKING + MERGE
    # ------------------------------------------------------------------
    log.info("Processando arquivo de Tracking...")
    df_trk = parse_excel_cello(TRCK_FILE)
    df_uni = pd.merge(df_trk, df_occ, on="DO No.", how="inner")

    # Classificar tipo de retorno
    df_uni["Return_type"] = np.where(df_uni["IOD Datetime"].notnull(), "Partial Return", "Full return")

    colunas = [
        "DO No.", "Order Sub Status", "Shipping Type", "DO Total Qty.",
        "Return Qty.", "NF Return No.", "IOD Datetime", "Occurrence Code",
        "Occurrence Name", "Shipment No.", "Return_type",
    ]
    # Verificar se todas as colunas existem
    colunas_faltando = [c for c in colunas if c not in df_uni.columns]
    if colunas_faltando:
        raise KeyError(f"Colunas faltando após merge: {colunas_faltando}")

    df_aux01 = df_uni[colunas].copy()
    df_aux01 = df_aux01[(df_aux01["IOD Datetime"].isna()) & (df_aux01["Return_type"] == "Full return")]
    df_aux01 = pd.merge(
        df_aux01, df_rtn_occ,
        left_on="Occurrence Code", right_on="occurrence", how="left"
    )
    log.info(f"{len(df_aux01)} registros para processar no Shipment.")
    log.info(f"\n{df_aux01}")

    # ------------------------------------------------------------------
    # PROCESSAR SHIPMENT — Full Return
    # ------------------------------------------------------------------
    log.info("Navegando para Shipment...")
    page.get_by_text("Transport Planning").nth(1).click()
    page.wait_for_timeout(500)
    page.get_by_role("paragraph").filter(has_text=re.compile(r"^Shipment$")).click(timeout=60000)

    iframe3 = page.frame_locator("iframe[id^=ifm_TMSS010000026557]")
    resultados = []

    for i, row in df_aux01.iterrows():
        do  = str(row["DO No."]).strip()
        rtn = row.get("rtn_s_reason_cd")  # coluna do merge com df_rtn_occ (por nome, não índice)
        log.info(f"[{i+1}/{len(df_aux01)}] DO: {do} | SAP reason: {rtn}")

        status_result = ""
        try:
            if pd.isna(rtn):
                log.warning(f"  Sub-reason não encontrado para DO {do}. Pulando.")
                status_result = "Return Sub-reason not found"
            else:
                rtn = str(rtn).strip()
                iframe3.locator("#ipt_doNo").click()
                iframe3.locator("#ipt_doNo").fill(do)
                page.wait_for_timeout(5000)
                iframe3.get_by_role("button", name="Search").click()
                iframe3.locator("#row0grd_shmpt").get_by_role("gridcell").filter(
                    has_text=re.compile(r"^$")
                ).first.click()

                iframe3.locator("#btn_refusalReturn").click()
                popup = page.locator("#iframePopup").content_frame

                # Selecionar opção FULL
                popup.locator("#dropdownlistArrowcmb_rtrnOpt_tmsShmptRtrnP02 div").click()
                page.wait_for_timeout(2000)
                popup.get_by_role("option", name="FULL").locator("span").click()

                # Digitar o reason code caractere por caractere (campo especial)
                popup.locator("#dropdownlistArrowcmb_rtrnRsnSap_tmsShmptRtrnP02 div").click()
                for char in rtn:
                    popup.locator("#cmb_rtrnRsnSap_tmsShmptRtrnP02").press(char)

                page.wait_for_timeout(5000)
                popup.locator("#cmb_rtrnRsnSap_tmsShmptRtrnP02").press("Enter")
                popup.get_by_role("columnheader").filter(has_text=re.compile(r"^$")).click()

                page.wait_for_timeout(15000)
                popup.get_by_role("link", name="close").click()

                status_result = "Refusal done"
                log.info(f"  Refusal concluído para DO {do}.")

        except TimeoutError:
            log.error(f"  Timeout ao processar DO {do}.")
            status_result = "Timeout error"
        except Exception as exc:
            log.exception(f"  Erro inesperado ao processar DO {do}: {exc}")
            status_result = f"Erro: {exc}"

        resultados.append({
            "DO No.":           do,
            "Order Sub Status": row.get("Order Sub Status"),
            "Shipping Type":    row.get("Shipping Type"),
            "DO Total Qty.":    row.get("DO Total Qty."),
            "Return Qty.":      row.get("Return Qty."),
            "NF Return No.":    row.get("NF Return No."),
            "Occurrence Code":  row.get("Occurrence Code"),
            "Occurrence Name":  row.get("Occurrence Name"),
            "Shipment No.":     row.get("Shipment No."),
            "Return_type":      row.get("Return_type"),
            "Return subreason": rtn if not pd.isna(rtn) else None,
            "Status":           status_result,
        })

    df_result = pd.DataFrame(resultados)
    log.info(f"\n{df_result}")

    # ------------------------------------------------------------------
    # SALVAR E ENVIAR E-MAIL
    # ------------------------------------------------------------------
    if not df_result.empty:
        df_result.to_excel(FULL_RTN_FILE, index=False)
        log.info(f"Resultado salvo em: {FULL_RTN_FILE}")

        tabela_html = df_result.to_html(index=False, justify="center")

        # CSS corrigido (faltavam chaves no tbody tr)
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
        <p>Segue abaixo e em anexo as DOs processadas.</p>
        {tabela_html}
        <p>Atenciosamente,</p>
        <p>RPA_full_return</p>
        </body>
        </html>
        """

        TO_ADDRESS = ["matheus.o@samsung.com"]
        # TO_ADDRESS = ["odair.jr@samsung.com", "egberto.s@samsung.com", "michael.hs@samsung.com"]
        CC_ADDRESS = []
        # CC_ADDRESS = ["carlos.ienne@samsung.com", "tarsis.r@samsung.com"]

        enviar_email(
            smtp_server    = "smtp.w2.samsung.net",
            smtp_port      = 25,
            from_addr      = email_from,
            password       = email_pass,
            to_list        = TO_ADDRESS,
            cc_list        = CC_ADDRESS,
            subject        = "[B2B] Full return process",
            body_html      = body_html,
            attachment_path= FULL_RTN_FILE,
        )
    else:
        log.info("Nenhum resultado para reportar.")

    log.info("RPA Full Return finalizado!")
    context.close()
    browser.close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    ts_inicio = datetime.now()
    log.info(f"Início: {ts_inicio.strftime('%Y-%m-%d %H:%M:%S')}")

    with sync_playwright() as playwright:
        run(playwright)

    ts_fim = datetime.now()
    log.info(f"Fim:    {ts_fim.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"Duração: {ts_fim - ts_inicio}")
