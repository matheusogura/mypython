# =============================================================================
# LLM CREATION SCRIPT — Versão Corrigida
# Correções aplicadas:
#   - Imports duplicados e não utilizados removidos
#   - fabio_maxuuu substituída por iteração direta de string
#   - fabão_nós_te_amamos definida UMA vez fora do loop, com parâmetros explícitos
#   - Verificação de mensagem2 corrigida (usava 'mensagem' por engano)
#   - all_email corrigida (lista aninhada → lista plana)
#   - Credenciais de e-mail corrigidas (os.getenv com nomes reais)
#   - Caminho de arquivo centralizado em constante
#   - pass após raise removido (código morto)
#   - Captura de exceções genéricas mais específica
#   - Nomes de variáveis e funções mais descritivos
#   - Valores numéricos nos dicts de caminhão declarados como float diretamente
#   - Detecção de popup "[Dynamic] CREATE_LLM" corrigida: uso de regex + wait_for_selector
#     em vez de get_by_text simples (que não esperava o popup aparecer)
# =============================================================================

import re
import os
import smtplib
from datetime import date, timedelta, datetime
from typing import List

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError

load_dotenv()

# ── Credenciais ────────────────────────────────────────────────────────────────
# CORREÇÃO: os.getenv('') sempre retornava None. Nomes de variável adicionados.
cello_id   = os.getenv('CELLO_ID')
cello_pass = os.getenv('CELLO_PASS')
email_from = os.getenv('EMAIL_FROM')
email_pass = os.getenv('EMAIL_PASS')

# ── Caminhos centralizados ─────────────────────────────────────────────────────
# CORREÇÃO: caminho hardcoded repetido 5x no código original. Agora é uma constante.
OUTPUT_DIR       = os.getenv('OUTPUT_DIR', 'C:/Users/matheus.o/Desktop/Downloads/')
TRACKING_FILE    = os.path.join(OUTPUT_DIR, 'LLM_creation.xlsx')
RESULT_FILE      = os.path.join(OUTPUT_DIR, 'LLMCREATE.xlsx')
LLM_OUTPUT_FILE  = os.path.join(OUTPUT_DIR, 'LLM.xlsx')

# ── Datas ──────────────────────────────────────────────────────────────────────
today     = date.today()
date_to   = '2026-02-16'
date_from = '2026-02-09'

# ── Tabelas de caminhões ───────────────────────────────────────────────────────
# CORREÇÃO: valores declarados como float diretamente, evitando 10x .astype(float)
norm_truck = {
    'min_weight':     [0.0, 620.0, 1200.0, 1800.0, 3000.0, 4000.0, 6000.0, 14000.0, 30000.0],
    'max_weight':     [620.0, 1200.0, 1800.0, 3000.0, 4000.0, 6000.0, 14000.0, 30000.0, 32000.0],
    'min_volume':     [0.0, 2.102815, 5.508, 8.16408, 14.244912, 24.2726, 36.363, 41.2114, 82.018064],
    'max_volume':     [2.102815, 5.508, 8.16408, 14.244912, 24.2726, 36.363, 41.2114, 82.018064, 87.216392],
    'truck_name':     ['TNT00.5','TNT01.3','TNT01.8','TNT02.4','TNT20.0','TNT25.0','TNT30.0','TNT40.0','TNT42.0'],
    'priority':       ['10','11','12','13','14','15','16','17','18'],
    'lenght':         ['173','240','290','368','590','690','782','1420','1510'],
    'width':          ['110','180','160','207','220','248','248','248','248'],
    'heiht':          ['110.5','127.5','175.95','187','187','212.5','212.5','232.9','232.9'],
    'load_possible_wt': [620.0, 1200.0, 1800.0, 3000.0, 4000.0, 6000.0, 14000.0, 30000.0, 32000.0],
    'volume':         [2.102815, 5.508, 8.16408, 14.244912, 24.2726, 36.363, 41.2114, 82.018064, 87.216392],
    'regra':          ['Normal']*9,
    'amount':         [7000000.0]*9,
}

prot_truck = {
    'min_weight':     [0.0, 2370.0, 5850.0, 10000.0],
    'max_weight':     [2370.0, 5850.0, 10000.0, 16000.0],
    'min_volume':     [0.0, 5.7185535, 22.728405, 35.44415],
    'max_volume':     [5.7185535, 22.728405, 35.44415, 81.98386],
    'truck_name':     ['TPT02.0','TPT04.0','TPT09.0','TPT25.0'],
    'priority':       ['20','21','22','23'],
    'lenght':         ['174','535','920','1480'],
    'width':          ['209','245','245','245'],
    'heiht':          ['157.25','173.4','157.25','226.1'],
    'load_possible_wt': [2370.0, 5850.0, 10000.0, 16000.0],
    'volume':         [5.7185535, 22.728405, 35.44415, 81.98386],
    'regra':          ['Protect']*4,
    'amount':         [45000000.0]*4,
}

n_trck = pd.DataFrame(norm_truck)
p_trck = pd.DataFrame(prot_truck)

# Colunas de prioridade e próximo tipo de caminhão (shift para pegar o próximo)
n_trck['new_priority']   = n_trck['priority'].shift(-1)
p_trck['new_priority']   = p_trck['priority'].shift(-1)
n_trck['new_truck_type'] = n_trck['truck_name'].shift(-1)
p_trck['new_truck_type'] = p_trck['truck_name'].shift(-1)


# =============================================================================
# FUNÇÃO: digitar_string_no_campo
# CORREÇÃO: a função original (fabio_maxuuu) criava uma lista de sub-listas de
# 1 caractere cada, percorrendo com while + índice manual. Em Python, strings
# já são iteráveis — basta um `for char in texto`. A função abaixo encapsula
# esse comportamento para reutilização, mas poderia ser inline também.
# =============================================================================
def digitar_string_no_campo(locator, texto: str):
    """Pressiona cada caractere de 'texto' no elemento Playwright fornecido."""
    for char in texto:
        locator.press(char)


# =============================================================================
# FUNÇÃO: mensagem_dinamica_apareceu
# NOVA CORREÇÃO: o get_by_text simples não esperava o popup aparecer, e o texto
# "[Dynamic] CREATE_LLM[MNF/..." é dinâmico (varia por manifesto). A solução usa
# wait_for_selector com regex para aguardar ativamente o elemento e fazer match
# parcial seguro, retornando True/False em vez de lançar exceção.
# =============================================================================
def mensagem_dinamica_apareceu(page, timeout_ms: int = 5000) -> bool:
    """
    Aguarda o popup de erro dinâmico do LLM aparecer na página.
    Usa regex para match parcial, pois o texto completo varia por manifesto
    (ex: "[Dynamic] CREATE_LLM[MNF/0001]").
    Retorna True se apareceu dentro do timeout, False caso contrário.
    """
    try:
        page.wait_for_selector("text=/\\[Dynamic\\] CREATE_LLM\\[MNF\\//", timeout=timeout_ms)
        return True
    except TimeoutError:
        return False


# =============================================================================
# FUNÇÃO: criar_manifesto_llm
# CORREÇÃO 1: a função era definida DUAS vezes — uma antes e uma dentro do loop,
#             a segunda sobrescrevendo a primeira a cada iteração.
# CORREÇÃO 2: usava variáveis externas (shtmt, df01) sem recebê-las como parâmetro,
#             o que tornava o comportamento frágil e difícil de testar.
# CORREÇÃO 4: a verificação de "[Dynamic] CREATE_LLM[MNF/" usava get_by_text simples,
#             que não esperava o popup aparecer e não fazia match com texto dinâmico.
#             Substituído por mensagem_dinamica_apareceu() com wait_for_selector + regex.
# =============================================================================
def criar_manifesto_llm(page, iframe, truck: str, truck2: str,
                         shtmt: str, df01, tab: list) -> list:
    """
    Seleciona o tipo de caminhão na interface e cria o manifesto LLM.
    - Para regra 'Normal':  tenta truck; só usa truck2 se a mensagem dinâmica aparecer.
    - Para regra 'Protect': mesmo comportamento — truck2 apenas se mensagem aparecer.
    truck2 NUNCA é escolhido automaticamente sem a mensagem dinâmica.
    Retorna a lista 'tab' atualizada com o resultado.
    """
    regra = df01['regra'].iloc[0] if 'regra' in df01.columns else 'Normal'

    def _selecionar_caminhao_na_lista(tipo_caminhao: str):
        """Digita o nome do caminhão na barra de scroll e clica nele."""
        scroll_bar = page.locator("#iframePopup").content_frame.locator(
            "#jqxScrollThumbverticalScrollBarinnerListBoxcmb_truckTcd_tmsMnf2ndCreateLLMP01"
        )
        # CORREÇÃO: substituído while + fabio_maxuuu por iteração direta
        for char in tipo_caminhao:
            scroll_bar.press(char)

        page.wait_for_timeout(500)
        elemento = page.locator("#iframePopup").content_frame.get_by_text(tipo_caminhao)

        # Rola para baixo até encontrar o elemento
        while elemento.count() == 0:
            try:
                page.locator("#iframePopup").content_frame.locator(
                    "#jqxScrollAreaDownverticalScrollBarinnerListBoxcmb_truckTcd_tmsMnf2ndCreateLLMP01"
                ).click()
                page.wait_for_timeout(500)
            except TimeoutError:
                # CORREÇÃO: except genérico substituído por TimeoutError específico
                page.wait_for_timeout(500)
                page.locator("#iframePopup").content_frame.get_by_text(tipo_caminhao, exact=True).click()

        page.wait_for_timeout(500)
        page.locator("#iframePopup").content_frame.get_by_text(tipo_caminhao, exact=True).click()

    def _salvar_e_confirmar():
        page.locator("#iframePopup").content_frame.locator(
            "#btn_save_tmsMnf2ndCreateLLMP01"
        ).click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="OK").click()
        page.wait_for_timeout(500)

    # ── Abertura do popup de criação ────────────────────────────────────────
    iframe.locator("#columntablegrd_smmy").get_by_role("columnheader").filter(
        has_text=re.compile(r"^$")
    ).click()
    iframe.locator("#btn_crtLlm01").click()

    # Seleciona regra "Ship-to Compatible"
    page.locator("#iframePopup").content_frame.locator(
        "#dropdownlistArrowcmb_llmRule_tmsMnf2ndCreateLLMP01 div"
    ).click()
    page.wait_for_timeout(500)
    page.locator("#iframePopup").content_frame.get_by_text("Ship-to Compatible").click()
    page.locator("#iframePopup").content_frame.locator(
        "#dropdownlistArrowcmb_llmRule_tmsMnf2ndCreateLLMP01 div"
    ).click()

    # Abre dropdown de tipo de caminhão e seleciona 'truck'
    page.locator("#iframePopup").content_frame.locator(
        "#dropdownlistArrowcmb_truckTcd_tmsMnf2ndCreateLLMP01 div"
    ).click()
    _selecionar_caminhao_na_lista(truck)
    _salvar_e_confirmar()

    # ── Verifica se houve erro de regra dinâmica ────────────────────────────
    # CORREÇÃO: truck2 (caminhão maior) só é usado se a mensagem dinâmica aparecer,
    # independente de ser Protect ou Normal. Sem a mensagem, sempre confirma com truck.
    if mensagem_dinamica_apareceu(page, timeout_ms=5000):
        print(f"Regra dinâmica ativada [{regra}] — tentando com caminhão alternativo: {truck2}")
        page.wait_for_timeout(1000)
        page.get_by_role("button", name="OK").click()
        page.wait_for_timeout(1000)

        # Tenta com truck2 — só chegamos aqui se a mensagem dinâmica apareceu
        page.locator("#iframePopup").content_frame.locator(
            "#dropdownlistArrowcmb_truckTcd_tmsMnf2ndCreateLLMP01 div"
        ).click()
        _selecionar_caminhao_na_lista(truck2)
        _salvar_e_confirmar()
        page.get_by_role("button", name="OK").click()

        msg_criado = page.get_by_text("1 Manifest(s) created.[")
        if msg_criado.is_visible() and "1 Manifest(s) created.[" in (msg_criado.inner_text() or ""):
            tab.append({'Shipment': shtmt, 'Truck': truck2,
                        'Carrier': df01['carrier_final'].iloc[0], 'Status': 'LLM Created - 1M'})
            print(f"Manifesto criado com truck2 [{truck2}] (1 manifesto)")
        else:
            tab.append({'Shipment': shtmt, 'Truck': truck2,
                        'Carrier': df01['carrier_final'].iloc[0], 'Status': 'LLM Created - +M'})

    else:
        # Mensagem dinâmica NÃO apareceu — confirma com truck original independente da regra
        print(f"Sem mensagem dinâmica [{regra}] — confirmando com truck original: {truck}")
        page.wait_for_timeout(1000)
        page.get_by_role("button", name="OK").click()

        msg_criado = page.get_by_text("1 Manifest(s) created.[")
        if msg_criado.is_visible() and "1 Manifest(s) created.[" in (msg_criado.inner_text() or ""):
            tab.append({'Shipment': shtmt, 'Truck': truck,
                        'Carrier': df01['carrier_final'].iloc[0], 'Status': 'LLM Created - 1M'})
            print(f"Manifesto criado com truck [{truck}] (1 manifesto)")
        else:
            tab.append({'Shipment': shtmt, 'Truck': truck,
                        'Carrier': df01['carrier_final'].iloc[0], 'Status': 'LLM Created - +M'})

    return tab


# =============================================================================
# PRIMEIRA EXECUÇÃO: login e limpeza de popup inicial
# =============================================================================
def run_login_inicial(playwright: Playwright) -> None:
    """Faz login e descarta o popup de 'Last Login Information' se aparecer."""
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page    = context.new_page()
    page.goto("http://105.202.220.4/cello/view/login.html?sso=N&idp=N", timeout=60000)

    page.get_by_role("textbox", name="User ID").fill(cello_id)
    page.get_by_role("textbox", name="User ID").press("Tab")
    page.get_by_role("textbox", name="Password").fill(cello_pass)
    page.get_by_role("button", name="LOGIN").click()

    mensagem = page.get_by_text("Last Login Information")
    if mensagem.is_visible() and "Last Login Information" in (mensagem.inner_text() or ""):
        print("Popup de login anterior detectado — fechando...")
        page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
        try:
            with page.expect_popup() as _:
                page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
        except TimeoutError:
            pass  # Popup secundário não apareceu — ok

    context.close()
    browser.close()


# =============================================================================
# SEGUNDA EXECUÇÃO: extração de dados e criação dos LLMs
# =============================================================================
def run_principal(playwright: Playwright) -> None:
    # DOCKER: headless=True obrigatório — container não tem display gráfico.
    # Para debugar localmente, mude HEADLESS=false no .env ou na linha de comando.
    headless = os.getenv('HEADLESS', 'true').lower() == 'true'
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context()
    page    = context.new_page()

    # ── Login ──────────────────────────────────────────────────────────────
    page.goto("http://105.202.220.4/cello/view/login.html?sso=N&idp=N#")
    page.get_by_role("textbox", name="User ID").fill(cello_id)
    page.get_by_role("textbox", name="User ID").press("Tab")
    page.get_by_role("textbox", name="Password").fill(cello_pass)
    page.get_by_role("button", name="LOGIN").click()

    mensagem = page.get_by_text("Last Login Information")
    if mensagem.is_visible() and "Last Login Information" in (mensagem.inner_text() or ""):
        page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
        try:
            with page.expect_popup() as _:
                page.locator("#modals-container").get_by_role("button", name="OK").click(timeout=5000)
        except TimeoutError:
            pass

    # ── Navegação até Transport Tracking ──────────────────────────────────
    page.locator("#btn_userGrpId").nth(1).click()
    page.get_by_text("SDSLA_SL-CAJ2").dblclick()
    page.locator("#menuTBox").get_by_text("TMS").click()
    page.locator("#sideMenu").get_by_text("Prime").click()
    page.locator("#sideMenu").get_by_text("Transport Tracking").click()
    page.get_by_role("paragraph").filter(has_text=re.compile(r"^Tracking Management$")).click()

    # ── Filtros na tela de Tracking ────────────────────────────────────────
    iframe = page.frame_locator("iframe[id^=ifm_TMSS010000034069]")

    # Status de envio
    iframe.locator("#dropdownlistArrowcmb_shmptScd div").click()
    page.wait_for_timeout(500)
    iframe.get_by_text("S400 : Ship", exact=True).click()
    iframe.get_by_text("S500 : Transit", exact=True).click()
    iframe.locator("#listitem5innerListBoxcmb_shmptScd").get_by_text("S600 : IOD").click()
    page.wait_for_timeout(500)

    iframe.locator("#ipt_doNo").fill("*")
    iframe.locator("#ipt_doNo").press("Tab")
    page.wait_for_timeout(500)
    iframe.get_by_role("button").filter(has_text=re.compile(r"^$")).nth(4).click()

    # Tipos de ordem
    tipos_ordem = [
        ("CAJ_IM_AIR",      "CAJ_IM_AIR : CAJ_IM_AIR"),
        ("CAJ_IM_TRUCK_LTL","CAJ_IM_TRUCK_LTL :"),
        ("MAO_IM_AIR",      "MAO_IM_AIR : MAO_IM_AIR"),
        ("MAO_IM_STO",      "MAO_IM_STO : MAO_IM_STO"),
    ]
    iframe.locator("#dropdownlistArrowcmb_ordTypTcd div").click()
    page.wait_for_timeout(500)
    for busca, texto_opcao in tipos_ordem:
        iframe.get_by_role("textbox", name="Looking for").fill(busca)
        iframe.get_by_role("textbox", name="Looking for").press('Enter')
        page.wait_for_timeout(500)
        iframe.locator("#listitem0innerListBoxcmb_ordTypTcd").get_by_text(texto_opcao).click()

    # Intervalo de datas
    # CORREÇÃO: substituído while + fabio_maxuuu por iteração direta (digitar_string_no_campo)
    campo_data_de = iframe.locator("#inputfromInput_dap_dlvryHopDate")
    campo_data_de.click()
    digitar_string_no_campo(campo_data_de, date_from)
    campo_data_de.press('Tab')

    campo_data_ate = iframe.locator("#inputtoInput_dap_dlvryHopDate")
    campo_data_ate.click()
    digitar_string_no_campo(campo_data_ate, date_to)
    campo_data_de.press('Tab')

    # Serviços detalhados
    iframe.locator("div").filter(has_text=re.compile(r"^Setting Reset Search$")).get_by_role("button").nth(3).click()
    iframe.locator("#dropdownlistArrowcmb_detlSvcTcd div").click()
    page.wait_for_timeout(500)
    iframe.get_by_text("C05 : LTL-AIR").click()

    servicos = [("t02", "T02 : LTNW"), ("CD", "CD : Truck-air")]
    for busca, texto_opcao in servicos:
        iframe.get_by_role("textbox", name="Looking for").click()
        page.wait_for_timeout(500)
        iframe.get_by_role("textbox", name="Looking for").fill(busca)
        iframe.get_by_role("textbox", name="Looking for").press("Enter")
        page.wait_for_timeout(500)
        iframe.locator("#listitem0innerListBoxcmb_detlSvcTcd").get_by_text(texto_opcao).click()
        page.wait_for_timeout(500)

    iframe.locator("#dropdownlistArrowcmb_detlSvcTcd div").click()
    iframe.get_by_role("button", name="Search").click()
    expect(
        iframe.locator("#row0grd_toList").get_by_role("gridcell").filter(has_text=re.compile(r"^$")).first
    ).to_be_visible(timeout=600000)

    # ── Download do Excel ──────────────────────────────────────────────────
    print(f"Iniciando download: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    max_tentativas = 10
    download = None

    for tentativa in range(1, max_tentativas + 1):
        try:
            with page.expect_download(timeout=600000) as download_info:
                iframe.locator("#btn_excelDown").click()
            download = download_info.value
            print(f"Download concluído na tentativa {tentativa}")
            break
        except Exception as e:
            print(f"Tentativa {tentativa}/{max_tentativas} falhou: {e}")
            page.wait_for_timeout(1500)
            try:
                page.get_by_role("button", name="OK").click()
            except TimeoutError:
                pass
            if tentativa == max_tentativas:
                print("Número máximo de tentativas atingido.")
                raise  # CORREÇÃO: 'pass' após 'raise' removido (código morto)

    download.save_as(TRACKING_FILE)

    try:
        page.get_by_role("button", name="OK").click()
    except TimeoutError:
        pass

    print(f"Download finalizado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # ── Processamento do Excel ─────────────────────────────────────────────
    page.locator("#Tabs-wrap").get_by_role("list").locator("div").nth(1).click()

    df = pd.read_excel(TRACKING_FILE)
    df.drop(index=[0, 1], inplace=True)
    df = df.reset_index(drop=True)
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    coluna_a_dropar = df.columns[0]
    df = df.drop(columns=[coluna_a_dropar])
    df = df.reset_index(drop=True)

    df['DO Total Vol.']  = df['DO Total Vol.'].astype(float)
    df['DO Total Wt']    = df['DO Total Wt'].astype(float)
    df['Invoice Amount'] = df['Invoice Amount'].astype(float)

    # Filtra apenas linhas sem manifesto de última perna
    df = df[df['Last Leg Manifest'].isna()]

    # Monta tabela de referência de caminhões
    df_carrier = pd.concat([p_trck, n_trck])

    # Extrai nome limpo da transportadora (remove código entre parênteses)
    df['carrier']     = df['Carrier Name'].str.split("(")
    df['carriername'] = df['carrier'].apply(
        lambda x: x[0] if isinstance(x, list) and len(x) > 0 else None
    )

    # ── Agrupamento por valor de invoice ──────────────────────────────────
    especiais = ['IBL VALORES ']
    df_amount = df.groupby(
        ['Arr Customer DC','Ship-to Name','carriername','UDF'], as_index=False
    )['Invoice Amount'].sum().reset_index(drop=True)

    df_amount['regra'] = df_amount['carriername'].apply(
        lambda x: "Protect" if x in especiais else 'Normal'
    )
    df_amount['pgr'] = df_amount['regra'].apply(
        lambda x: 45_000_000 if x == 'Protect' else 7_000_000
    )
    df_amount['qty_trucks'] = np.ceil(df_amount['Invoice Amount'] / df_amount['pgr'])

    # ── Agrupamento por shipment ───────────────────────────────────────────
    chave_grupo = ['Arr Customer DC','Ship-to Name','carriername','UDF']
    df_01 = df.groupby(
        chave_grupo + ['Shipment No.'], as_index=False
    ).agg({'DO Total Vol.': 'sum', 'DO Total Wt': 'sum', 'Invoice Amount': 'sum'})

    df_01 = pd.merge(
        df_01,
        df_amount[chave_grupo + ['regra','pgr','qty_trucks']],
        how='left', on=chave_grupo
    )

    # ── Shipments que precisam de 1 caminhão (consolidados) ───────────────
    df_cons = df_01[df_01['qty_trucks'] == 1].reset_index(drop=True)

    # ── Shipments que precisam de 2 caminhões ─────────────────────────────
    df2 = df_01[df_01['qty_trucks'] == 2].sort_values(
        by=['Arr Customer DC','Ship-to Name','UDF','carriername','regra','Invoice Amount']
    ).reset_index(drop=True)

    # Distribui os shipments de 2 caminhões em manifestos respeitando o limite de valor
    manifestos = []
    manifesto_atual = 1
    soma_atual = 0.0

    for _, row in df2.iterrows():
        limite = row['pgr']
        if soma_atual + row['Invoice Amount'] > limite:
            manifesto_atual += 1
            soma_atual = row['Invoice Amount']
        else:
            soma_atual += row['Invoice Amount']
        manifestos.append(str(manifesto_atual))

    df2['Manifesto'] = manifestos

    df_02 = df2.groupby(
        ['Manifesto'] + chave_grupo + ['regra']
    ).agg({'DO Total Vol.': 'sum', 'DO Total Wt': 'sum', 'Invoice Amount': 'sum'}).reset_index()

    # ── Matching de tipo de caminhão por volume e peso ─────────────────────
    def _match_caminhao(df_alvo: pd.DataFrame) -> pd.DataFrame:
        """
        Faz merge_asof por volume e peso para determinar o tipo de caminhão adequado,
        depois seleciona o de maior prioridade entre os dois critérios.
        """
        df_carrier_sorted_vol = df_carrier.sort_values('min_volume').reset_index(drop=True)
        df_carrier_sorted_wt  = df_carrier.sort_values('min_weight').reset_index(drop=True)

        df_alvo = df_alvo.sort_values('DO Total Vol.').reset_index(drop=True)
        df_alvo = pd.merge_asof(
            df_alvo,
            df_carrier_sorted_vol[['min_volume','max_volume','truck_name','priority','regra']],
            left_on='DO Total Vol.', right_on='min_volume',
            by='regra', direction='backward'
        )

        df_alvo = df_alvo.sort_values('DO Total Wt').reset_index(drop=True)
        df_alvo = pd.merge_asof(
            df_alvo,
            df_carrier_sorted_wt[['min_weight','max_weight','truck_name','priority','regra']],
            left_on='DO Total Wt', right_on='min_weight',
            by='regra', direction='backward'
        )

        df_alvo['Considerado_priority'] = df_alvo[['priority_x','priority_y']].max(axis=1)

        df_alvo = pd.merge(
            df_alvo,
            df_carrier[['truck_name','priority']],
            how='left', left_on='Considerado_priority', right_on='priority'
        ).rename(columns={'priority': 'priority_left'})

        df_alvo = pd.merge(
            df_alvo,
            df_carrier[['new_truck_type','priority','new_priority']],
            how='left', left_on='Considerado_priority', right_on='priority'
        )
        return df_alvo

    if not df_02.empty:
        df_02 = _match_caminhao(df_02)

    # ── Consolidados: agrupamento e match de caminhão ──────────────────────
    df_sub_vol = df_cons.groupby(chave_grupo + ['regra'], as_index=False)['DO Total Vol.'].sum()
    df_sub_wt  = df_cons.groupby(chave_grupo + ['regra'], as_index=False)['DO Total Wt'].sum()

    df_carrier_sorted_vol = df_carrier.sort_values('min_volume').reset_index(drop=True)
    df_carrier_sorted_wt  = df_carrier.sort_values('min_weight').reset_index(drop=True)

    df_result_vol = pd.merge_asof(
        df_sub_vol.sort_values('DO Total Vol.').reset_index(drop=True),
        df_carrier_sorted_vol[['min_volume','truck_name','priority','regra']],
        left_on='DO Total Vol.', right_on='min_volume',
        by='regra', direction='backward'
    )

    df_result_wt = pd.merge_asof(
        df_sub_wt.sort_values('DO Total Wt').reset_index(drop=True),
        df_carrier_sorted_wt[['min_weight','max_weight','truck_name','priority','regra']],
        left_on='DO Total Wt', right_on='min_weight',
        by='regra', direction='backward'
    )

    df_completed = pd.merge(
        df_result_vol, df_result_wt,
        how='left', on=chave_grupo
    )
    df_completed['Considerado_priority'] = df_completed[['priority_x','priority_y']].max(axis=1)
    df_completed = pd.merge(
        df_completed,
        df_carrier[['truck_name','priority']],
        how='left', left_on='Considerado_priority', right_on='priority'
    ).rename(columns={'priority': 'priority_left'})
    df_completed = pd.merge(
        df_completed,
        df_carrier[['new_truck_type','priority','new_priority']],
        how='left', left_on='Considerado_priority', right_on='priority'
    ).reset_index()

    # ── Dados auxiliares de shipment ───────────────────────────────────────
    df_aux = df.groupby(chave_grupo + ['Shipment No.']).agg(
        {'DO Total Vol.': 'sum', 'DO Total Wt': 'sum', 'Invoice Amount': 'sum'}
    ).reset_index()
    df_aux['total_amount'] = df_aux.groupby(chave_grupo)['Invoice Amount'].transform('sum')

    # ── Combinação final ───────────────────────────────────────────────────
    if not df_02.empty:
        df_03 = pd.concat([df_completed, df_02], axis=0).reset_index(drop=True)
        df_03 = df_03.reset_index()
        df_03['Grupo'] = df_03['index']
        df_03 = df_03.drop_duplicates()

        df_03 = pd.merge(
            df_03,
            df2[chave_grupo + ['Manifesto','Shipment No.']],
            how='left', on=chave_grupo + ['Manifesto']
        )
        df_03 = pd.merge(
            df_03,
            df[chave_grupo + ['Shipment No.']],
            on=chave_grupo, how='left', suffixes=('', '_new')
        )
        df_03['Shipment No.'] = df_03['Shipment No.'].fillna(df_03['Shipment No._new'])
        df_03.drop(columns=['Shipment No._new'], inplace=True)
        df_03 = df_03.drop_duplicates()
    else:
        df_03 = df_completed.reset_index()
        df_03['Grupo'] = df_03['index']
        df_03 = pd.merge(
            df_03,
            df[chave_grupo + ['Shipment No.']],
            how='left', on=chave_grupo
        )

    df_03 = pd.merge(df_03, df_aux, how='left', on=chave_grupo + ['Shipment No.'])
    df_03 = df_03[[
        'Arr Customer DC','Ship-to Name','carriername','UDF',
        'DO Total Vol._y','min_volume','truck_name_x',
        'DO Total Wt_y','min_weight','max_weight','truck_name_y',
        'Considerado_priority','truck_name','priority_left','new_truck_type',
        'Grupo','Shipment No.','Invoice Amount',
    ]].drop_duplicates().reset_index(drop=True)

    # CORREÇÃO: 'regra' vinha como regra_x/regra_y dependendo dos merges e era
    # removida na seleção do df_03. Agora é trazida diretamente do df_amount,
    # que tem a fonte original limpa sem sufixos.
    df_03 = pd.merge(
        df_03,
        df_amount[chave_grupo + ['regra']].drop_duplicates(),
        how='left',
        on=chave_grupo
    )

    # ── Grupo final com transportadora e destino ───────────────────────────
    grupo = df_03[[
        'Arr Customer DC','Ship-to Name','carriername','UDF','Grupo',
        'Shipment No.','truck_name','DO Total Vol._y','DO Total Wt_y',
        'Invoice Amount','new_truck_type','regra'
    ]].reset_index(drop=True)

    grupo['carrier_final'] = np.where(grupo['UDF'].isin(['MS','MT']), 'UNIDOCKS', grupo['carriername'])
    grupo['resultado']     = np.where(grupo['carriername'] == grupo['carrier_final'], 'others', 'DHL')
    grupo.to_excel(RESULT_FILE, index=False)

    # ── Criação dos LLMs na interface ─────────────────────────────────────
    tab = []
    page.get_by_text("Transport Planning").nth(1).click()
    page.get_by_text("Create LLM").click()
    iframe = page.frame_locator("iframe[id^=ifm_TMSS010000034295]")

    print(f"Total de grupos a processar: {grupo['Grupo'].nunique()}")

    for nome_grupo, dados_grupo in grupo.groupby('Grupo'):
        df01 = grupo[grupo['Grupo'] == nome_grupo]

        # Lista de shipments do grupo
        lista_shipments = df01['Shipment No.'].tolist()
        shtmt = ",".join(str(s) for s in lista_shipments)

        # Tipo de caminhão primário e alternativo
        truck  = df01['truck_name'].drop_duplicates().iloc[0]
        truck2 = df01['new_truck_type'].drop_duplicates().iloc[0]

        if not lista_shipments:
            continue  # Grupo vazio — pular

        page.wait_for_timeout(500)

        # Configura opções de carrier e porto
        iframe.locator("#dropdownlistArrowcmb_carrierOpt div").click()
        page.wait_for_timeout(500)
        iframe.get_by_role("option", name="1st Carrier").locator("span").click()
        iframe.locator("#dropdownlistWrappercmb_hubPortOpt").click()
        page.wait_for_timeout(500)
        iframe.get_by_role("option", name="Dest. Port").locator("span").click()

        print(f"Processando grupo {nome_grupo}: {shtmt}")

        # Preenche campo de número de shipment
        page.wait_for_timeout(500)
        iframe.locator("#btn_shipmentNo").click()
        page.wait_for_timeout(500)
        iframe.locator("textarea").fill(shtmt)
        page.wait_for_timeout(500)
        iframe.get_by_role("button", name="Apply").click()
        iframe.get_by_text("Search").click()
        page.wait_for_timeout(1000)

        test2 = iframe.locator("#row0grd_smmy").get_by_role("gridcell").filter(
            has_text=re.compile(r"^$")
        ).locator("div").nth(2)

        eh_dhl = (df01['resultado'] == 'DHL').any()

        # Define sequência de opções de porto a tentar
        opcoes_porto = (
            [("2nd Carrier", "Dest. Port"), ("1st Carrier", "Arr. Port"), ("1st Carrier", "Hub")]
            if eh_dhl
            else [("1st Carrier", "Dest. Port"), ("1st Carrier", "Arr. Port"), ("1st Carrier", "Hub")]
        )

        manifesto_criado = False
        try:
            for carrier_opt, porto_opt in opcoes_porto:
                iframe.locator("#dropdownlistArrowcmb_carrierOpt div").click()
                page.wait_for_timeout(500)

                if carrier_opt == "2nd Carrier":
                    iframe.get_by_text("2nd Carrier").click()
                else:
                    iframe.get_by_role("option", name="1st Carrier").locator("span").click()

                iframe.locator("#dropdownlistArrowcmb_hubPortOpt div").click()
                page.wait_for_timeout(500)

                if porto_opt == "Hub":
                    iframe.get_by_text("Hub", exact=True).click()
                elif porto_opt == "Arr. Port":
                    iframe.get_by_text("Arr. Port").click()
                else:
                    iframe.get_by_role("option", name="Dest. Port").locator("span").click()

                iframe.get_by_text("Search").click()
                page.wait_for_timeout(1000)

                test2 = iframe.locator("#row0grd_smmy").get_by_role("gridcell").filter(
                    has_text=re.compile(r"^$")
                ).locator("div").nth(2)

                if test2.count() > 0:
                    # CORREÇÃO: função agora definida fora do loop e recebe parâmetros explícitos
                    tab = criar_manifesto_llm(page, iframe, truck, truck2, shtmt, df01, tab)
                    manifesto_criado = True
                    break

            if not manifesto_criado:
                tab.append({
                    'Shipment': shtmt, 'Truck': truck,
                    'Carrier': df01['carrier_final'].iloc[0],
                    'Status': 'Shipment not found'
                })

        except TimeoutError:
            tab.append({
                'Shipment': shtmt, 'Truck': truck,
                'Carrier': df01['carrier_final'].iloc[0],
                'Status': 'Timeout - Shipment not found'
            })

    # ── Salva resultado e envia e-mail ─────────────────────────────────────
    tabela = pd.DataFrame(tab).reset_index(drop=True)
    print(tabela)
    tabela.to_excel(LLM_OUTPUT_FILE, index=False)

    tabela_html = tabela.to_html(index=False, justify='center')

    corpo_email = f"""
    <html>
    <style>
        .styled-table {{
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 0.9em;
            font-family: sans-serif;
            min-width: 400px;
            box-shadow: 0 0 20px rgba(0,0,0,0.15);
        }}
        .styled-table th, .styled-table td {{ padding: 12px 15px; }}
        .styled-table tbody tr:nth-of-type(even) {{ background-color: #f3f3f3; }}
        .styled-table tbody tr:last-of-type {{ border-bottom: 2px solid #009879; }}
    </style>
    <body>
        <p>Olá!</p>
        <p>Segue abaixo os LLMs criados.</p>
        {tabela_html}
        <p>Atenciosamente,</p>
        <p>RPA</p>
    </body>
    </html>"""

    SMTP_SERVER = "smtp.w2.samsung.net"
    SMTP_PORT   = 25
    TO_ADDRESS  = [
        'matheus.o@samsung.com',
        'felipe.fm@samsung.com',
        'fabio.mfc@samsung.com',
        'gabrielle.so@partner.samsung.com',
    ]
    CC_ADDRESS  = []

    msg = MIMEMultipart()
    msg['From']    = email_from
    msg['To']      = ",".join(TO_ADDRESS)
    msg['CC']      = ",".join(CC_ADDRESS)
    msg['Subject'] = 'LLM CREATION'

    msg_alt = MIMEMultipart('alternative')
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText("Visualize em HTML para melhor formatação.", 'plain'))
    msg_alt.attach(MIMEText(corpo_email, 'html'))

    # CORREÇÃO: nome do anexo era o caminho completo. Agora usa apenas o nome do arquivo.
    nome_anexo = os.path.basename(LLM_OUTPUT_FILE)
    with open(LLM_OUTPUT_FILE, 'rb') as f:
        attachment = MIMEApplication(f.read(), Name=nome_anexo)
        attachment['Content-Disposition'] = f'attachment; filename="{nome_anexo}"'
        msg.attach(attachment)

    # CORREÇÃO: all_email era [TO+CC] (lista dentro de lista).
    # sendmail aceita lista diretamente — sem loop necessário.
    todos_destinatarios = TO_ADDRESS + CC_ADDRESS

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.login(email_from, email_pass)
            server.sendmail(email_from, todos_destinatarios, msg.as_string())
        print(f"E-mail enviado para {todos_destinatarios}")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

    context.close()
    browser.close()


# =============================================================================
# PONTO DE ENTRADA
# =============================================================================
if __name__ == "__main__":
    inicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"Início: {inicio}")

    with sync_playwright() as pw:
        run_login_inicial(pw)

    with sync_playwright() as pw:
        run_principal(pw)

    fim = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"Fim: {fim}")
