"""
RPA - Automação de E-mail, Excel e Sistema
==========================================
Orquestrador : Prefect 3.x
Agendamento  : Seg–Sex | 09:00–23:59 | a cada 10 ou 30 minutos
Autor        : <seu nome>
"""
import re
import ssl
import os
import poplib
import email
import glob
import smtplib
import logging
from email import policy
from email.message import EmailMessage
from datetime import datetime, date, timedelta
from pathlib import Path
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError
import xlwings as xw
import pandas as pd
from prefect import flow, task, get_run_logger
from prefect.client.schemas.schedules import CronSchedule
from dotenv import load_dotenv
from typing import Dict
from tnefparse import TNEF
import json


# ---------------------------------------------------------------------------
# CONFIGURAÇÕES  –  ajuste conforme seu ambiente
# ---------------------------------------------------------------------------
load_dotenv()
POP3_HOST     = os.getenv("POP3_SERVER")
POP3_PORT     = int(os.getenv("POP3_PORT"))
POP3_USER     = os.getenv("email_in_mat")
POP3_PASSWORD = os.getenv("pass_mail_mat")

USER        = os.getenv("email_in_mat")
PASSWORD    = os.getenv("pass_mail_mat")
POP3_SERVER = "pop3.w2.samsung.net"
PORT        = 995        # POP3 + SSL
POP3_USE_SSL  = True

SMTP_HOST     = os.getenv("SMTP_SERVER")
SMTP_PORT     = os.getenv("SMTP_PORT")
SMTP_USER     = os.getenv("email_in_mat")
SMTP_PASSWORD = os.getenv("pass_mail_mat")

USER_CELLO    = os.getenv('cello_id')
USER_PASS     = os.getenv('cello_pass')

# Palavras-chave aceitas no assunto (case-insensitive, basta uma coincidir)
SUBJECT_KEYWORDS = ['antecipação','antecipacão','antecipacao','antecipaçao']

# Quantidade máxima de e-mails mais recentes a varrer por execução.
# O POP3 não tem busca server-side, então limitar o range acelera a varredura.
# Ajuste conforme o volume da caixa: 100 é seguro para caixas com fluxo moderado.
MAX_EMAILS_TO_SCAN = 100

# Filtro de remetentes desativado a pedido da operação.
# Para reativar, descomente a lista abaixo, a função _sender_matches()
# e o bloco `if not _sender_matches(msg)` dentro de buscar_emails_pop3.
# ALLOWED_SENDERS = [
#     'remetente1@samsung.com',
#     'remetente2@samsung.com',
# ]

# Pasta onde o anexo será salvo (limpa antes de cada execução)
DOWNLOAD_FOLDER = Path(os.environ.get("DOWNLOAD_FOLDER", r"C:/Users/matheus.o/Desktop/Downloads/attachments"))
CELLO_FILE       = os.path.join(DOWNLOAD_FOLDER, 'to_list_adiant.xlsx')

# Nome fixo que será dado ao arquivo baixado
ATTACHMENT_FILENAME = "input_data.xlsx"
EXTENSIONS_OK = {".xlsx", ".xls", ".xlsm"}

EXCEL_COLUMN    = "A"
EXCEL_START_ROW = 2

# Destinatários do relatório final
REPORT_RECIPIENTS = ['matheus.o@samsung.com',
    'emeneghel@samsung.com',
    'meire.as@samsung.com',
    'andrew.f@samsung.com',
    'fabio.mfc@samsung.com',
    'felipe.fm@samsung.com',
    'andrea.lopes@samsung.com',
    'gabrielle.so@partner.samsung.com',
    'v.almeida@samsung.com',
    'camila.mf@partner.samsung.com',
    'marcio.rm@partner.samsung.com',
    'c.damiao@partner.samsung.com',
    ]
REPORT_SUBJECT    = f"Relatório RPA – {date.today().strftime('%d/%m/%Y')}"
TOLIST = "tolist_ant.xlsx"

data = datetime.now()


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def digitar_string_no_campo(locator, texto: str):
    """Pressiona cada caractere de 'texto' no elemento Playwright fornecido."""
    for char in texto:
        locator.press(char)

def _subject_matches(subject: str) -> bool:
    """Verifica se o assunto contém alguma palavra-chave configurada."""
    subject_lower = subject.lower()
    return any(kw.lower() in subject_lower for kw in SUBJECT_KEYWORDS)


# Filtro de remetentes desativado a pedido da operação — mantido comentado para uso futuro.
# def _sender_matches(msg: email.message.Message) -> bool:
#     """Verifica se o remetente do e-mail está na lista de remetentes permitidos."""
#     from_header = msg.get("From", "")
#     from_lower = from_header.lower()
#     return any(sender.lower() in from_lower for sender in ALLOWED_SENDERS)


def _save_excel(payload: bytes, dest: Path, original_name: str) -> None:
    """Grava o payload em disco e faz log."""
    dest.write_bytes(payload)
    logger = get_run_logger()
    logger.info(f"Anexo Excel salvo em {dest} (nome original: '{original_name}')")


def _extrair_de_winmail(dat_bytes: bytes) -> list[tuple[bytes, str]]:
    """
    Recebe o conteúdo binário de um winmail.dat (TNEF) e devolve
    uma lista de tuplas (payload, filename) para todos os anexos encontrados.
    """
    logger = get_run_logger()
    try:
        tnef = TNEF(dat_bytes)
    except Exception as exc:
        logger.error(f"Falha ao abrir winmail.dat como TNEF: {exc}")
        return []

    anexos = []
    for att in tnef.attachments:
        nome = att.name or f"anexo_tnef_{len(anexos)}"
        anexos.append((att.data, nome))
        logger.debug(f"Encontrado no winmail.dat → {nome} ({len(att.data)} bytes)")
    return anexos


def arrumar_excel_pandas(CELLO_FILE):
    """
    Lê o arquivo Excel e limpa linhas completamente vazias no início,
    promove o primeiro registro não-vazio como cabeçalho e remove a primeira coluna.
    """
    df = pd.read_excel(CELLO_FILE)
    df.drop(index=[0, 1], inplace=True) # Remove as linhas com índices 0 e 1 - Linhas vazias
    df = df.reset_index(drop=True)
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    coluna_a_dropar = df.columns[0]
    df = df.drop(columns=[coluna_a_dropar])
    return df

def _get_email_date(msg: email.message.Message) -> date | None:
    """Extrai a data do cabeçalho 'Date' do e-mail."""
    raw_date = msg.get("Date", "")
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
        "%d %b %Y %H:%M:%S %z",
    ):
        try:
            return datetime.strptime(raw_date.strip(), fmt).date()
        except ValueError:
            continue
    return None


def save_attachment(part, folder: Path) -> Path:
    filename = part.get_filename()
    if not filename:
        raise ValueError("Attachment has no filename")
    safe_name = normalise_name(filename)
    dest = folder / safe_name
    dest.parent.mkdir(parents=True, exist_ok=True)
    data = part.get_payload(decode=True)
    dest.write_bytes(data)
    return dest


def normalise_name(name: str) -> str:
    return " ".join(name.strip().split()).lower()


def has_valid_ext(name: str) -> bool:
    _, ext = os.path.splitext(name)
    return ext.lower() in EXTENSIONS_OK


def _normalizar_valor(v):
    """
    Converte valores lidos do Excel para o formato correto antes de enviar ao sistema.

    - float "inteiro" (ex.: 12345678.0) → int (12345678), evitando a notação '.0'
    - float com casas decimais reais (ex.: 12.5) → mantém como float
    - outros tipos (str, None) → retorna sem alterar
    """
    if isinstance(v, float):
        # Se é um float que representa um inteiro, converte para int
        if v.is_integer():
            return int(v)
    return v


def _ler_excel_helper(arquivo: Path) -> list:
    """
    Abre o arquivo Excel com xlwings e extrai os valores da coluna
    configurada em EXCEL_COLUMN a partir da linha EXCEL_START_ROW.
    Separado como helper puro para reuso sem passar pelo sistema de tasks do Prefect.

    Normaliza automaticamente valores float que representam inteiros
    (ex.: 12345678.0 → 12345678), evitando o sufixo '.0' quando o valor
    for convertido para string.
    """
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(arquivo))
        ws = wb.sheets[0]
        celula_inicio = f"{EXCEL_COLUMN}{EXCEL_START_ROW}"
        intervalo = ws.range(celula_inicio).expand("down")
        valores = intervalo.value
        wb.close()
    finally:
        app.quit()

    if not isinstance(valores, list):
        valores = [valores]

    # Filtra nulos e normaliza cada valor (float inteiro → int)
    return [_normalizar_valor(v) for v in valores if v is not None]


# ---------------------------------------------------------------------------
# TASK 1 – Leitura do e-mail via POP3
# ---------------------------------------------------------------------------

def _get_email_datetime(msg: email.message.Message) -> datetime | None:
    """Extrai a data e hora do cabeçalho 'Date' do e-mail (para ordenação cronológica)."""
    raw_date = msg.get("Date", "")
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
        "%d %b %Y %H:%M:%S %z",
    ):
        try:
            return datetime.strptime(raw_date.strip(), fmt)
        except ValueError:
            continue
    return None


@task(name="buscar_emails_pop3", retries=2, retry_delay_seconds=10)
def buscar_emails_pop3() -> list[dict]:
    """
    Conecta ao servidor POP3 e retorna TODOS os e-mails do dia que:
      - tenham assunto com alguma palavra-chave de SUBJECT_KEYWORDS
      - UID ainda não registrado em processados.txt
      (filtro de remetente está desativado — ver ALLOWED_SENDERS comentado no topo)

    Lista é devolvida em ordem cronológica (mais antigo → mais novo).
    Retorna lista vazia se não houver e-mail válido.
    """
    logger = get_run_logger()
    ssl_ctx = ssl.create_default_context()

    if POP3_USE_SSL:
        conn = poplib.POP3_SSL(POP3_HOST, POP3_PORT, context=ssl_ctx)
    else:
        conn = poplib.POP3(POP3_HOST, POP3_PORT)
    print("\n🔗 Conectando ao servidor POP3...")

    try:
        conn.user(POP3_USER)
        conn.pass_(POP3_PASSWORD)
    except poplib.error_proto as e:
        print("Erro ao autenticar:", e)
        raise
    print("✅ Autenticação concluída.\n")

    num_messages = len(conn.list()[1])
    logger.info(f"Caixa de entrada: {num_messages} mensagem(ns)")

    hoje = date.today()

    # Carrega UIDs já processados
    uid_file = Path("processados.txt")
    processados = set(uid_file.read_text().splitlines()) if uid_file.exists() else set()

    # OTIMIZAÇÃO 1: pega o mapa completo de UIDs em uma única chamada
    # em vez de uma chamada por mensagem (economiza round-trips de rede).
    uidl_response = conn.uidl()
    uid_map: dict[int, str] = {}
    for line in uidl_response[1]:
        parts = line.decode().split()
        if len(parts) >= 2:
            uid_map[int(parts[0])] = parts[1]

    encontrados: list[dict] = []

    # OTIMIZAÇÃO 2: varre apenas os N e-mails mais recentes.
    # Como o RPA roda a cada 30 min em horário comercial, MAX_EMAILS_TO_SCAN
    # cobre folgadamente o que chegou no intervalo.
    start_idx = max(1, num_messages - MAX_EMAILS_TO_SCAN + 1)
    logger.info(f"Varrendo mensagens {start_idx} a {num_messages} "
                f"(últimas {num_messages - start_idx + 1}).")

    for i in range(num_messages, start_idx - 1, -1):

        # OTIMIZAÇÃO 3: checa UID processado ANTES de baixar qualquer coisa.
        uid = uid_map.get(i)
        if uid is None:
            logger.warning(f"UID não encontrado para a mensagem {i}.")
            continue

        if uid in processados:
            # E-mail já processado em execução anterior — pula sem baixar.
            continue

        # OTIMIZAÇÃO 4: usa TOP para baixar APENAS os cabeçalhos (zero linhas do corpo).
        # Muito mais rápido que RETR, que traz a mensagem inteira com anexos.
        try:
            header_lines = conn.top(i, 0)[1]
        except poplib.error_proto as exc:
            logger.warning(f"Falha em TOP {i}: {exc} — pulando.")
            continue

        header_bytes = b"\n".join(header_lines)
        header_msg = email.message_from_bytes(header_bytes, policy=policy.default)

        # Filtra data (usando só os cabeçalhos)
        data_email = _get_email_date(header_msg)
        if data_email != hoje:
            continue

        # Filtro de remetente desativado a pedido da operação.
        # if not _sender_matches(header_msg):
        #     continue

        # Filtra assunto (usando só os cabeçalhos)
        assunto = header_msg.get("Subject", "")
        if not _subject_matches(assunto):
            continue

        # Passou em todos os filtros — AGORA sim baixa a mensagem completa
        # com anexos para processar depois.
        logger.info(f"Match encontrado (msg {i}, UID {uid}) — baixando mensagem completa.")
        raw = b"\n".join(conn.retr(i)[1])
        msg = email.message_from_bytes(raw, policy=policy.default)

        # Datetime para ordenação cronológica
        dt = _get_email_datetime(msg) or datetime.min

        encontrados.append({
            "msg": msg,
            "uid": uid,
            "datetime": dt,
            "assunto": assunto,
            "from": msg.get("From", ""),
        })

    conn.quit()

    # Ordena do mais antigo para o mais novo
    encontrados.sort(key=lambda item: item["datetime"])

    if not encontrados:
        logger.info("Nenhum e-mail novo e válido encontrado para hoje.")
    else:
        logger.info(f"{len(encontrados)} e-mail(s) novo(s) encontrado(s) para processar:")
        for i, e in enumerate(encontrados, 1):
            logger.info(f"  [{i}] {e['datetime'].strftime('%H:%M:%S')} | {e['from']} | UID={e['uid']}")

    return encontrados


# ---------------------------------------------------------------------------
# TASK 2 – Download do anexo
# ---------------------------------------------------------------------------

@task(name="baixar_anexo")
def baixar_anexo(email_data: Dict) -> Path:
    """
    Busca o primeiro anexo Excel (.xlsx/.xls/.xlsm) na mensagem.
    Se o e‑mail contiver um winmail.dat, tenta extrair anexos dele.
    """
    logger = get_run_logger()

    DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    for f in DOWNLOAD_FOLDER.iterdir():
        f.unlink()
    logger.debug(f"Pasta de download limpa: {DOWNLOAD_FOLDER}")

    msg: email.message.EmailMessage = email_data["msg"]
    uid: str = email_data["uid"]
    destino = DOWNLOAD_FOLDER / ATTACHMENT_FILENAME

    encontrado = False

    # Primeiro tenta os anexos "normais"
    for part in msg.iter_attachments():
        filename = part.get_filename() or f"anexo_{uid}.bin"
        ext = os.path.splitext(filename.lower())[1]
        logger.debug(f"Anexo detectado → filename={filename!r}, ext={ext!r}")

        if ext not in EXTENSIONS_OK:
            continue

        payload = part.get_payload(decode=True)
        if payload:
            _save_excel(payload, destino, filename)
            encontrado = True
            break

        logger.warning(f"Payload vazio para {filename!r}")

    # Se ainda não achou, verifica se há winmail.dat
    if not encontrado:
        for part in msg.iter_parts():
            ctype = part.get_content_type()
            if ctype == "application/ms-tnef":
                logger.info("Detectado winmail.dat – tentando extrair TNEF...")
                winmail_bytes = part.get_payload(decode=True)
                if not winmail_bytes:
                    logger.warning("Payload do winmail.dat está vazio.")
                    continue

                for payload, nome in _extrair_de_winmail(winmail_bytes):
                    ext = os.path.splitext(nome.lower())[1]
                    if ext in EXTENSIONS_OK:
                        _save_excel(payload, destino, nome)
                        encontrado = True
                        break
                if encontrado:
                    break

    if not encontrado:
        logger.error("Nenhum anexo Excel encontrado após analisar winmail.dat.")
        for i, p in enumerate(msg.iter_parts()):
            logger.error(
                f"[Parte {i}] type={p.get_content_type()}, "
                f"disp={p.get('Content-Disposition')}, "
                f"filename={p.get_filename()}"
            )
        raise FileNotFoundError(
            "Nenhum anexo .xlsx/.xls/.xlsm encontrado no e‑mail. "
            "Confira o log acima para a listagem das partes da mensagem."
        )

    return destino


@task(name="marcar_uid_processado")
def marcar_uid_processado(uid: str) -> None:
    """
    Registra o UID como processado em processados.txt.
    Chamado APÓS o processamento bem-sucedido do e-mail.
    """
    logger = get_run_logger()
    uid_file = Path("processados.txt")
    uid_file.parent.mkdir(parents=True, exist_ok=True)
    with uid_file.open("a", encoding="utf-8") as f:
        f.write(uid + "\n")
    logger.info(f"UID {uid} registrado como processado.")


# ---------------------------------------------------------------------------
# TASK 3 – Leitura do Excel com xlwings
# ---------------------------------------------------------------------------

@task(name="ler_excel_xlwings")
def ler_excel_xlwings(arquivo: Path) -> list:
    """
    Abre o arquivo Excel com xlwings e extrai os valores da coluna
    configurada em EXCEL_COLUMN a partir da linha EXCEL_START_ROW.
    Retorna uma lista com os valores não nulos.
    """
    logger = get_run_logger()
    logger.info(f"Abrindo Excel: {arquivo}")
    lista = _ler_excel_helper(arquivo)
    logger.info(f"Coluna {EXCEL_COLUMN} – {len(lista)} item(ns) encontrado(s)")
    return lista


# ---------------------------------------------------------------------------
# TASK 4 – Automação do sistema (Playwright)
# ---------------------------------------------------------------------------

@task(name="automatizar_sistema")
def automatizar_sistema(lista: list) -> pd.DataFrame:
    """
    Itera sobre a lista extraída do Excel e interage com o sistema
    via Playwright.
    """
    logger = get_run_logger()
    logger.info(f"Iniciando automação para {len(lista)} item(ns)...")
    df_aux = pd.DataFrame(lista, columns=["do no."])
    
    df_aux = df_aux['do no.'].astype(str)
    # CORRIGIDO: usa o parâmetro `lista` recebido pela task
    # em vez de reler o arquivo com ler_excel_xlwings() internamente
    # dos = json.dumps(lista)
    lista = json.dumps(lista)
    
    lista = lista.replace('"', '')
    lista = lista.replace('[', '')
    lista = lista.replace(']', '')
    # logger.info(f"Lista: {lista}")
    # TODO: resultados ainda não são populados — implementar coleta de retorno
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Faz login
        page.goto("http://105.202.220.4/cello/view/login.html?sso=N&idp=N#", timeout=60000)
        page.get_by_role("textbox", name="User ID").fill(USER_CELLO)
        page.get_by_role("textbox", name="User ID").press("Tab")
        page.get_by_role("textbox", name="Password").fill('Marco@!03')
        page.get_by_role("button", name="LOGIN").click()

        # Trata popup de login anterior
        mensagem = page.get_by_text("Last Login Information")
        if mensagem.is_visible() and "Last Login Information" in (mensagem.inner_text() or ""):
            print("Popup de login anterior detectado — fechando...")
            page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
            try:
                with page.expect_popup() as _:
                    page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
            except TimeoutError:
                pass

        page.locator("#btn_userGrpId").nth(1).click()
        page.get_by_text("SDSLA_SL-CAJ2").dblclick()
        page.wait_for_timeout(500)

        page.locator("#menuTBox").get_by_text("TMS").click()
        page.locator("#sideMenu").get_by_text("Prime").click()
        page.bring_to_front()
        # Acessa TO List
        page.get_by_text("Transport Order").nth(1).click()
        page.get_by_role("paragraph").filter(has_text=re.compile(r"^T/O List$")).click()
        page.get_by_text("Transport Order").nth(1).click()

        ifm = page.frame_locator("iframe[id^=ifm_TMSS01000002655]")

        # ifm.locator("#dropdownlistContentcmb_stsCombo1").click()
        # page.wait_for_timeout(500)
        # ifm.locator("#jqxScrollAreaDownverticalScrollBarinnerListBoxcmb_stsCombo1").click()
        # ifm.locator("#jqxScrollAreaDownverticalScrollBarinnerListBoxcmb_stsCombo1").click()
        # page.wait_for_timeout(500)
        # ifm.locator("#listitem5innerListBoxcmb_stsCombo1").get_by_text("D085 : Booking Confirmed(").click()
        # ifm.locator("#listitem4innerListBoxcmb_stsCombo1").get_by_text("D080 : Booking Confirmed").click()

        ifm.locator("#btn_doNo").click()
        ifm.locator("#multiInputipt_doNo").get_by_role("textbox").click()
        ifm.locator("#multiInputipt_doNo").get_by_role("textbox").fill(lista)
        ifm.get_by_role("button", name="Apply").click()


        ifm.get_by_role("button", name="Search").click()

        expect(ifm.get_by_role("gridcell").first).to_be_visible()

        # Download do arquivo Excel
        with page.expect_download() as download_info:
            ifm.locator("#btn_excelDown").click()
        download = download_info.value
        download.save_as(CELLO_FILE)

        # Processa o arquivo Excel
        df = arrumar_excel_pandas(CELLO_FILE)
        
        # -- Tratamento do dataframe ------------------------------
        df['PGI Date'] = pd.to_datetime(df['PGI Date'])
        df1 = df
        # # df1 = df[df['PGI Date'] <= data]
        df1 = df1[df1['Shipment No.'].isna()]
        # df1 = df1[df1['Sub'].str.strip().isin(['Booking Confirmed', 'Booking Confirmed(Block)'])]

        # df1 = df1[(df1['Sub']=='Booking Confirmed')|(df1['Sub']=='Booking Confirmed(Block)')]
        # df1 = df.head(15) # Apenas para apresentação / TESTE
        # logger.info(df)
        df1.to_csv('C:/Users/matheus.o/Desktop/Downloads/attachments/dos.csv')
        df2 = df1['Del note'].to_json(orient='records') 
        # df3 = df.iloc[0,7]
        df2 = df2.replace('"', '')
        df2 = df2.replace('[', '')
        df2 = df2.replace(']', '')
        logger.info(f"df2: {df2}")
        

        # Acessa Internal Screen
        page.get_by_text("Transport Planning").nth(1).click()
        page.get_by_role("paragraph").filter(has_text="Internal Screen").click()

        iframe = page.frame_locator("iframe[id^=ifm_TMSS010000028695]")
        iframe.get_by_role("tab", name="T/O").click()
        resultados = []

        page.wait_for_timeout(500)
        iframe.locator("#btn_doNo").click()
        page.wait_for_timeout(500)
        iframe.locator("textarea").click()
        page.wait_for_timeout(500)
        iframe.locator("textarea").fill(df2)
        # page.wait_for_timeout(1500)
        iframe.get_by_role("button", name="Apply").click()
        # iframe.get_by_role("button", name="Search").click()
        iframe.get_by_text("Search").click()

        page.wait_for_timeout(500)
        
        texto = 'D920 : Shipped'
        campo = iframe.locator("#dropdownlistContentcmb_doSubStsTcd")
                
        campo.click()
        page.wait_for_timeout(500)
        digitar_string_no_campo(campo, texto)
        page.wait_for_timeout(500)
        campo.get_by_text(texto).click()
        iframe.locator("#dropdownlistContentcmb_doStsTcd").click()
        page.wait_for_timeout(500)
        iframe.get_by_text("D900 : Assigned").first.click()
        # page.wait_for_timeout(10000)
        # -- aplicar salvar -----------------------------------------------
        iframe.locator("#btn_doApply").click()
        
        page.wait_for_timeout(500)
        iframe.locator("#btn_doSave").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="OK").click() 
        logger.info("Aplicando salvar...")

        # Acessa Tracking Management
        page.locator("#sideMenu").get_by_text("Transport Tracking").click()
        page.get_by_role("paragraph").filter(has_text=re.compile(r"^Tracking Management$")).click()
        ifm1 = page.frame_locator("iframe[id^=ifm_TMSS010000034069]")

        ifm1.locator("#dropdownlistArrowcmb_shmptScd div").click()
        page.wait_for_timeout(500)
        ifm1.get_by_text("S400 : Ship", exact=True).click()
        ifm1.get_by_text("S500 : Transit", exact=True).click()
        ifm1.locator("#listitem5innerListBoxcmb_shmptScd").get_by_text("S600 : IOD").click()
        page.wait_for_timeout(500)

        ifm1.locator("#ipt_doNo").fill(df2)
        ifm1.get_by_role("button", name="Search").click()
        page.wait_for_timeout(500)

        expect(ifm1.locator("#row0grd_toList").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first).to_be_visible(timeout=600000)
        page.wait_for_timeout(1000)
        # ifm1.locator("#columntablegrd_toList").get_by_role("columnheader").filter(has_text=re.compile(r"^$")).click()
        # page.wait_for_timeout(500)
        ifm1.locator("#columntablegrd_toList").get_by_role("checkbox").click()
        page.wait_for_timeout(1000)
        ifm1.locator("#btn_futhProc").click()

        # -- inserir Next Step (No Action) --------------------------------------
        campo1 =page.locator("#iframePopup").content_frame.locator("#dropdownlistContentcmb_occur_tmsFuthProcP01")
        texto1 = '(B2B) Antecipa'
        campo1.click()
        page.wait_for_timeout(1000)        
        digitar_string_no_campo(campo1, texto1)    
        campo1.get_by_text(texto1).click()

        page.locator("#iframePopup").content_frame.locator("#dropdownlistContentcmb_nextStep_tmsFuthProcP01").click()
        page.wait_for_timeout(500)
        page.locator("#iframePopup").content_frame.get_by_text("No Action").click()

        # -- inserir Billing Occurrence (N/A) --------------------------------------
        campo2 = page.locator("#iframePopup").content_frame.locator("#dropdownlistContentcmb_billOccur_tmsFuthProcP01")
        campo2.click()
        page.wait_for_timeout(500)
        texto2 = 'N/A'
        digitar_string_no_campo(campo2, texto2)  
        campo2.get_by_text(texto2).click()
        campo2.press('Tab')

        # -- salvar a ocorrencia ---------------------------------------------
        page.locator("#iframePopup").content_frame.get_by_text("Save").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="OK").click(timeout = 120000)
        # page.locator("#iframePopup").content_frame.get_by_role("link", name="close").click() # Apenas para apresentação / TESTE
        # page.wait_for_timeout(5000)

        # -- Volta para Internal Screen --------------------------------------
        page.locator("#Tabs-wrap").get_by_text("Internal Screen").click(timeout = 240000)
        page.wait_for_timeout(500)
        iframe.get_by_text('Search').click()
        page.wait_for_timeout(500)
        # page.locator("#Tabs-wrap").get_by_text("Internal Screen").click()
        df_aux1 = df1[(df1['Sub']=='Booking Confirmed')]
        if not df_aux1.empty: 
            df4 = df_aux1['Del note'].to_json(orient='records') 
            df4 = df4.replace('"', '')
            df4 = df4.replace('[', '')
            df4 = df4.replace(']', '')   
            iframe.locator("#btn_doNo").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").fill(df4)
            iframe.get_by_role("button", name="Apply").click()
            iframe.get_by_text("Search").click()

            texto2 = 'D080 : Booking Confirmed'
            campo = iframe.locator("#dropdownlistContentcmb_doSubStsTcd")
            try:
                campo.click()
                iframe.get_by_text(texto2).click()
            except Exception:
                # CORRIGIDO: usa keyboard.type() em vez de iterar caractere a caractere
                campo.click()
                page.keyboard.type(texto2)
                campo.press('Tab')
                       
            iframe.locator("#dropdownlistContentcmb_doStsTcd").click()
            iframe.locator("#listitem0innerListBoxcmb_doStsTcd").get_by_text("D000 : Hold").click()
                        # -- aplicar salvar -----------------------------------------------
            iframe.locator("#btn_doApply").click()
            
            page.wait_for_timeout(500)
            iframe.locator("#btn_doSave").click()
            page.wait_for_timeout(500)
            page.get_by_role("button", name="OK").click() 
        else:
            pass

        df_aux2 = df1[(df1['Sub']=='Booking Confirmed(Block)')]
            
        if not df_aux2.empty:
            df4 = df_aux2['Del note'].to_json(orient='records') 
            df4 = df4.replace('"', '')
            df4 = df4.replace('[', '')
            df4 = df4.replace(']', '')   
            iframe.locator("#btn_doNo").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").fill(df4)
            iframe.get_by_role("button", name="Apply").click()
            iframe.get_by_text("Search").click()

            texto2 = 'D085 : Booking Confirmed(Block)'
            campo = iframe.locator("#dropdownlistContentcmb_doSubStsTcd")
            try:
                campo.click()
                iframe.get_by_text(texto2).click()
            except Exception:
                # CORRIGIDO: usa keyboard.type() em vez de iterar caractere a caractere
                campo.click()
                page.keyboard.type(texto2)
                campo.press('Tab')
                       
            iframe.locator("#dropdownlistContentcmb_doStsTcd").click()
            iframe.locator("#listitem0innerListBoxcmb_doStsTcd").get_by_text("D000 : Hold").click()  
            # -- aplicar salvar -----------------------------------------------
            iframe.locator("#btn_doApply").click()
            
            page.wait_for_timeout(500)
            iframe.locator("#btn_doSave").click()
            page.wait_for_timeout(500)
            page.get_by_role("button", name="OK").click() 
        else:
            pass

        df_aux2 = df1[(df1['Sub']=='DO block')]
            
        if not df_aux2.empty:
            df4 = df_aux2['Del note'].to_json(orient='records') 
            df4 = df4.replace('"', '')
            df4 = df4.replace('[', '')
            df4 = df4.replace(']', '')   
            iframe.locator("#btn_doNo").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").click()
            page.wait_for_timeout(500)
            iframe.locator("textarea").fill(df4)
            iframe.get_by_role("button", name="Apply").click()
            iframe.get_by_text("Search").click()

            texto2 = 'D025 : DO block'
            campo = iframe.locator("#dropdownlistContentcmb_doSubStsTcd")
            try:
                campo.click()
                iframe.get_by_text(texto2).click()
            except Exception:
                # CORRIGIDO: usa keyboard.type() em vez de iterar caractere a caractere
                campo.click()
                page.keyboard.type(texto2)
                campo.press('Tab')
                       
            iframe.locator("#dropdownlistContentcmb_doStsTcd").click()
            iframe.locator("#listitem0innerListBoxcmb_doStsTcd").get_by_text("D000 : Hold").click()  
            # -- aplicar salvar -----------------------------------------------
            iframe.locator("#btn_doApply").click()
            
            page.wait_for_timeout(500)
            iframe.locator("#btn_doSave").click()
            page.wait_for_timeout(500)
            page.get_by_role("button", name="OK").click() 
        else:
            pass    
        # logger.info("Aplicando salvar... parte 2")

        browser.close()
        df_final = pd.merge(
            df_aux,
            df1[['Del note']],
            how = 'left',
            left_on = 'do no.',
            right_on = 'Del note',
        )
        df_aux = df_final[['do no.','Del note']]
        df_aux['Status'] = df_aux['Del note'].apply(lambda x: 'Checar delivery' if pd.isna(x) else 'Ok')
        resultados = df_aux[['do no.','Status']]
        df = pd.DataFrame(resultados)
        
        return pd.DataFrame(resultados)
    # return pd.DataFrame(resultados)


# ---------------------------------------------------------------------------
# TASK 5 – Envio do relatório por e-mail
# ---------------------------------------------------------------------------

@task(name="enviar_relatorio")
def enviar_relatorio(df: pd.DataFrame, resumo_lote: list[dict] | None = None) -> None:
    """
    Salva o DataFrame consolidado como Excel e envia por e-mail via SMTP.
    resumo_lote (opcional): lista com dicts descrevendo cada e-mail processado no lote,
    usada para montar uma seção no corpo do e-mail.
    """
    logger = get_run_logger()

    relatorio_path = DOWNLOAD_FOLDER / f"relatorio_{date.today().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(relatorio_path, index=False)
    logger.info(f"Relatório salvo: {relatorio_path}")

    # Monta seção de resumo do lote, se fornecida
    resumo_html = ""
    if resumo_lote:
        linhas = "".join(
            f"<tr><td>{i}</td><td>{e.get('datetime_str', '')}</td>"
            f"<td>{e.get('from', '')}</td><td>{e.get('uid', '')}</td>"
            f"<td>{e.get('itens', 0)}</td><td>{e.get('status', '')}</td></tr>"
            for i, e in enumerate(resumo_lote, 1)
        )
        resumo_html = f"""
        <p><b>E-mails processados neste lote:</b></p>
        <table border="1" cellpadding="4" cellspacing="0">
          <thead>
            <tr><th>#</th><th>Recebido</th><th>Remetente</th><th>UID</th><th>Itens</th><th>Status</th></tr>
          </thead>
          <tbody>{linhas}</tbody>
        </table>
        """

    corpo_html = f"""
    <html><body>
    <p>Olá,</p>
    <p>Segue em anexo o relatório consolidado de processamento de {date.today().strftime('%d/%m/%Y')}.</p>
    <p>Total de itens processados: {len(df)}</p>
    {resumo_html}
    <p>Segue abaixo a tabela com os resultados:</p>
    {df.to_html(index=False)}
    <p>Atenciosamente,<br>RPA Automação</p>
    </body></html>
    """

    msg = EmailMessage()
    msg["Subject"] = REPORT_SUBJECT
    msg["From"]    = SMTP_USER
    msg["To"]      = ", ".join(REPORT_RECIPIENTS)
    msg.set_content("Olá,\n\nSegue em anexo o relatório de processamento.\n\nAtenciosamente,\nRPA Automação")
    msg.add_alternative(corpo_html, subtype="html")

    with open(relatorio_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=relatorio_path.name,
        )

    with smtplib.SMTP(SMTP_HOST, int(SMTP_PORT)) as smtp:
        # smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)

    logger.info(f"Relatório enviado para: {REPORT_RECIPIENTS}")


# ---------------------------------------------------------------------------
# FLOW PRINCIPAL
# ---------------------------------------------------------------------------

@flow(name="rpa_antecipation_flow", log_prints=True)
def rpa_antecipation() -> None:
    """
    Flow principal do RPA em modo LOTE.

    Comportamento:
      1. Busca TODOS os e-mails novos do dia que batem com os filtros
         (assunto + remetentes + UID ainda não processado).
      2. Processa um de cada vez, em ordem cronológica (mais antigo primeiro).
      3. Após cada sucesso, marca o UID como processado.
      4. Se algum e-mail falhar, interrompe o lote (os que já foram processados
         ficam marcados; os restantes serão tentados na próxima execução).
      5. No final, envia UM relatório consolidado com os resultados de todos
         os e-mails processados com sucesso no lote.
    """
    logger = get_run_logger()
    logger.info("=" * 60)
    logger.info(f"Execução iniciada: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    # 1. Busca TODOS os e-mails válidos do dia
    emails = buscar_emails_pop3()
    if not emails:
        logger.info("Flow encerrado: sem e-mails válidos para processar.")
        return

    logger.info(f"Iniciando processamento em lote de {len(emails)} e-mail(s).")

    dfs_resultados: list[pd.DataFrame] = []
    resumo_lote: list[dict] = []
    erro_lote: Exception | None = None

    # 2. Processa cada e-mail sequencialmente
    for idx, email_data in enumerate(emails, 1):
        uid = email_data["uid"]
        remetente = email_data["from"]
        dt_str = email_data["datetime"].strftime("%d/%m/%Y %H:%M:%S")

        logger.info("-" * 60)
        logger.info(f"[{idx}/{len(emails)}] Processando e-mail "
                    f"recebido em {dt_str} | de: {remetente} | UID: {uid}")

        try:
            # 2.1 Download do anexo
            arquivo_excel = baixar_anexo(email_data)

            # 2.2 Leitura do Excel
            lista = ler_excel_xlwings(arquivo_excel)
            if not lista:
                logger.warning(f"[{idx}/{len(emails)}] Planilha sem dados — pulando este e-mail.")
                resumo_lote.append({
                    "datetime_str": dt_str, "from": remetente, "uid": uid,
                    "itens": 0, "status": "Sem dados na planilha",
                })
                # Marca como processado mesmo assim, para não reler o mesmo e-mail vazio
                marcar_uid_processado(uid)
                continue

            # 2.3 Automação do sistema (Playwright)
            df_resultados = automatizar_sistema(lista)
            dfs_resultados.append(df_resultados)

            # 2.4 Sucesso — registra UID como processado
            marcar_uid_processado(uid)

            resumo_lote.append({
                "datetime_str": dt_str, "from": remetente, "uid": uid,
                "itens": len(lista), "status": "OK",
            })
            logger.info(f"[{idx}/{len(emails)}] Concluído com sucesso ({len(lista)} item(ns)).")

        except Exception as exc:
            logger.error(f"[{idx}/{len(emails)}] Falha ao processar e-mail UID {uid}: {exc}")
            resumo_lote.append({
                "datetime_str": dt_str, "from": remetente, "uid": uid,
                "itens": 0, "status": f"ERRO: {exc}",
            })
            erro_lote = exc
            # Opção (a): interrompe o lote. E-mails restantes ficam para a próxima execução.
            break

    # 3. Envio do relatório consolidado (mesmo com erro, envia o que conseguiu)
    if dfs_resultados:
        df_consolidado = pd.concat(dfs_resultados, ignore_index=True)
        enviar_relatorio(df_consolidado, resumo_lote)
    elif resumo_lote:
        # Não houve resultados bem-sucedidos, mas há o que reportar (ex.: todos falharam)
        df_vazio = pd.DataFrame(columns=["Observação"])
        df_vazio.loc[0] = ["Nenhum e-mail foi processado com sucesso. Veja o resumo acima."]
        enviar_relatorio(df_vazio, resumo_lote)

    # 4. Se houve erro, propaga no final para o Prefect marcar o flow como falho
    if erro_lote is not None:
        logger.error("Flow encerrado com erro. E-mails restantes serão tentados na próxima execução.")
        raise erro_lote

    logger.info("Flow concluído com sucesso.")
    logger.info("=" * 60)


# ---------------------------------------------------------------------------
# PONTO DE ENTRADA – agendamento Prefect
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # rpa_antecipation.serve(
    #     name="rpa-antecipation_flow",
    #     schedules=[
    #         CronSchedule(
    #             cron="0/10 8-23 * * 1-5",
    #             timezone="America/Sao_Paulo",
    #         ),
    #         # A cada 10 min, 00h, Seg-Sex (Cobre a madrugada/final da noite)
    #         CronSchedule(
    #             cron="0/10 0 * * 1-5",                
    #             timezone="America/Sao_Paulo",
    #         )
    #     ],
    # )
    rpa_antecipation()
