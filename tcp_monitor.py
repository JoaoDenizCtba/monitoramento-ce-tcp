import argparse
import asyncio
import json
import msvcrt
import os
import sys
import tempfile
import time
import traceback
from datetime import datetime
from io import BytesIO
from contextlib import contextmanager, nullcontext, redirect_stderr, redirect_stdout
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from playwright.async_api import Error as PlaywrightError
from playwright.async_api import async_playwright
from pywintypes import com_error
import win32event

from excel_open import (
    EdicaoConcorrente, ExcelIndisponivel, aplicar_alteracoes, copiar_planilha, encontrar_planilha,
    planilha_por_ciclo,
)
from monitor_control import controles, enviar


URL = "https://portal.tcp.com.br/publico/v1/aduaneiro/consulta-ce-mercante"
API_PATH = "/legado/Aduaneiro/api/CeMercantePresencaImportacao"
TIMEOUT = 30000
MONITOR_HEADERS = ["CE", "Monitorar", "Última consulta", "Situação da consulta"]
PORTAL_HEADERS = [
    "CARGA SOLTA", "INVENTÁRIO", "CE", "CONTÊINER", "BL ORIGINAL",
    "DATA PRESENÇA", "DATA DESCARGA", "VIAGEM", "PORTO ORIGEM",
    "PORTO DESTINO", "AVARIA DO CONTÊINER", "LACRE", "DESOVADO",
    "DATA ESTUFAGEM", "VISTORIA SIGVIG", "ATUALIZAÇÃO MAPA",
    "REPRESENTANTE SIGVIG",
]
RESULT_HEADERS = ["ÚLTIMA CONSULTA", "MENSAGEM DO PORTAL"] + PORTAL_HEADERS
LEGACY_RESULT_HEADERS = PORTAL_HEADERS + ["MENSAGEM DO PORTAL", "ÚLTIMA CONSULTA"]
CONFIG_SHEET = "Configurações"
CONFIG_HEADERS = ["Configuração", "Valor"]
INTERVALS = (5, 15, 30, 45, 60)
SELECAO_PENDENTE = "Selecione Sim ou Não"


async def consultar(page, numero_ce):
    response = await page.goto(URL, wait_until="domcontentloaded", timeout=TIMEOUT)
    if response is None or not response.ok:
        raise RuntimeError(f"Falha ao abrir o portal: HTTP {response.status if response else 'sem resposta'}")
    frame = page.frame_locator('iframe[src*="/PortalV2/aduaneiro/consulta-ce-mercante"]')
    await frame.locator("#numero-ce").fill(numero_ce)
    async with page.expect_response(
        lambda r: API_PATH.lower() in r.url.lower() and r.request.method == "GET",
        timeout=TIMEOUT,
    ) as consulta:
        await frame.locator("#btn").click()
    resposta = await consulta.value
    if not resposta.ok:
        raise RuntimeError(f"Consulta rejeitada: HTTP {resposta.status}")
    dados = await resposta.json()
    if not isinstance(dados, dict) or dados.get("Status") != 1:
        raise RuntimeError("Resposta de consulta inesperada.")
    registros = dados.get("Objeto")
    if not isinstance(registros, list) or dados.get("TotalRegistros") != len(registros):
        raise RuntimeError("Lista ou quantidade de resultados inválida.")

    linhas = []
    if registros:
        rows = frame.locator('table tbody tr[ng-repeat-start="row in grid.Data.Objeto"]')
        await rows.first.wait_for(state="visible", timeout=TIMEOUT)
        tabela = rows.first.locator("xpath=ancestor::table[1]")
        colunas = [" ".join(x.split()) for x in await tabela.locator("thead th").all_inner_texts()]
        if colunas != PORTAL_HEADERS:
            raise RuntimeError("As colunas do portal não correspondem à planilha.")
        for row in await rows.all():
            cells = [" ".join(x.split()) for x in await row.locator("td").all_inner_texts()]
            if len(cells) != len(colunas):
                raise RuntimeError("Quantidade de células não corresponde às colunas.")
            if cells[2] != numero_ce:
                raise RuntimeError("O CE retornado pelo portal não corresponde ao consultado.")
            linhas.append(cells)
        if len(linhas) != len(registros):
            raise RuntimeError("A tabela não apresenta todos os registros da consulta.")
        linhas.sort()
    else:
        await frame.locator(".sem-resultado").wait_for(state="visible", timeout=TIMEOUT)

    mensagens = []
    for label in await frame.locator("form fieldset label:visible").all():
        texto = " ".join((await label.inner_text()).split())
        if texto.startswith(f"CE {numero_ce} "):
            mensagens.append(texto)
    return {"mensagens": sorted(mensagens), "linhas": linhas}


def selecionar_ces(aba):
    tabela = aba.tables.get("TabelaMonitoramento")
    cabecalho = 1
    if tabela is not None:
        primeira_coluna, cabecalho, ultima_coluna, _ = range_boundaries(tabela.ref)
        if primeira_coluna != 1 or ultima_coluna != 4:
            raise ValueError("A tabela de monitoramento deve ocupar as colunas A até D.")
    if [aba.cell(cabecalho, i).value for i in range(1, 5)] != MONITOR_HEADERS:
        raise ValueError("Os cabeçalhos da aba Monitoramento não são os esperados.")
    selecionados = []
    vistos = set()
    alterou = False
    for row in range(cabecalho + 1, aba.max_row + 1):
        ce = aba.cell(row, 1).value
        flag = aba.cell(row, 2).value
        situacao = aba.cell(row, 4)
        if flag is None or str(flag).strip() == "":
            if ce is not None and str(ce).strip():
                print(f"Linha {row}: selecione Sim ou Não em Monitorar; CE não consultado.")
                if situacao.value != SELECAO_PENDENTE:
                    situacao.value = SELECAO_PENDENTE
                    alterou = True
            elif situacao.value == SELECAO_PENDENTE:
                situacao.value = None
                alterou = True
            continue
        if situacao.value == SELECAO_PENDENTE:
            situacao.value = None
            alterou = True
        if str(flag).strip().casefold() in ("não", "nao"):
            continue
        if str(flag).strip().casefold() != "sim":
            raise ValueError(f"Valor inválido em Monitorar na linha {row}: {flag!r}.")
        if not isinstance(ce, str) or len(ce) != 15 or not ce.isascii() or not ce.isdigit():
            raise ValueError(f"O CE na linha {row} deve ser texto com exatamente 15 dígitos.")
        if ce in vistos:
            raise ValueError(f"O CE na linha {row} está duplicado entre os itens monitorados.")
        vistos.add(ce)
        selecionados.append((row, ce))
    return selecionados, alterou


def criar_aba_configuracoes(workbook):
    aba = workbook.create_sheet(CONFIG_SHEET)
    aba.append(CONFIG_HEADERS)
    aba.append(["Intervalo (minutos)", 60])
    aba.append(["Automação ativa", "Não"])
    aba.freeze_panes = "A2"
    aba.column_dimensions["A"].width = 32
    aba.column_dimensions["B"].width = 20
    for cell in aba[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for cell in (aba["B2"], aba["B3"]):
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.alignment = Alignment(horizontal="center")
    aba["B2"].number_format = "0"
    aba["B2"].comment = Comment("Escolha o intervalo entre consultas automáticas.", "Copilot")
    aba["B3"].comment = Comment("Sim ativa o modo contínuo; Não mantém apenas consultas manuais.", "Copilot")
    interval_validation = DataValidation(type="list", formula1='"5,15,30,45,60"', allow_blank=False)
    active_validation = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=False)
    for validation in (interval_validation, active_validation):
        validation.showErrorMessage = True
        validation.errorStyle = "stop"
        validation.errorTitle = "Seleção inválida"
        validation.error = "Escolha uma opção da lista."
    aba.add_data_validation(interval_validation)
    aba.add_data_validation(active_validation)
    interval_validation.add(aba["B2"])
    active_validation.add(aba["B3"])
    for row in aba.iter_rows(min_row=1, max_row=3, max_col=2):
        for cell in row:
            cell.protection = Protection(locked=cell.coordinate not in ("B2", "B3"))
    aba.protection.sheet = True
    return aba


def preparar_configuracoes(workbook):
    if CONFIG_SHEET not in workbook or "Monitoramento" not in workbook:
        raise ValueError("As abas Configurações e Monitoramento são necessárias.")
    aba = workbook[CONFIG_SHEET]
    if [aba.cell(1, col).value for col in (1, 2)] != CONFIG_HEADERS:
        raise ValueError("Os cabeçalhos da aba Configurações não são os esperados.")
    if aba["A2"].value != "Intervalo (minutos)" or aba["A3"].value != "Automação ativa":
        raise ValueError("Os campos da aba Configurações não são os esperados.")
    if aba.protection.sheet:
        raise ValueError("A aba Configurações já está protegida; nenhuma alteração foi feita.")
    intervalos = [v for v in aba.data_validations.dataValidation if "B2" in v.sqref]
    flags = [v for v in aba.data_validations.dataValidation if "B3" in v.sqref]
    if (len(intervalos) != 1 or intervalos[0].type != "list"
            or intervalos[0].formula1 not in ('"15,30,45,60"', '"5,15,30,45,60"')
            or len(flags) != 1 or flags[0].type != "list"
            or flags[0].formula1 != '"Sim,Não"'):
        raise ValueError("As listas de configuração não correspondem ao modelo esperado.")
    monitor_flags = [
        v for v in workbook["Monitoramento"].data_validations.dataValidation
        if "B2" in v.sqref and v.type == "list" and v.formula1 == '"Sim,Não"'
    ]
    if len(monitor_flags) != 1:
        raise ValueError("A lista Monitorar não corresponde ao modelo esperado.")
    intervalos[0].formula1 = '"5,15,30,45,60"'
    for validation in (intervalos[0], flags[0], monitor_flags[0]):
        validation.showErrorMessage = True
        validation.errorStyle = "stop"
        validation.errorTitle = "Seleção inválida"
        validation.error = "Escolha uma opção da lista."
    intervalos[0].allow_blank = False
    flags[0].allow_blank = False
    aba["B2"] = 5
    for row in aba.iter_rows(min_row=1, max_row=3, max_col=2):
        for cell in row:
            cell.protection = Protection(locked=cell.coordinate not in ("B2", "B3"))
    aba.protection.sheet = True


def validar_planilha_fechada(path):
    lock_path = path.with_name(f"~${path.name}")
    if lock_path.exists():
        raise RuntimeError("Feche a planilha no Excel antes de executar a automação.")


def garantir_aba_configuracoes(path, excel_book=None):
    if excel_book is not None:
        workbook = copiar_planilha(excel_book)
        try:
            if CONFIG_SHEET not in workbook:
                raise ValueError("A aba Configurações não existe; feche a planilha para criá-la.")
        finally:
            workbook.close()
        return
    validar_planilha_fechada(path)
    stat = path.stat()
    original_stat = (stat.st_mtime_ns, stat.st_size)
    workbook = load_workbook(path)
    try:
        if CONFIG_SHEET not in workbook.sheetnames:
            criar_aba_configuracoes(workbook)
            salvar_planilha(workbook, path, original_stat)
    finally:
        workbook.close()


def ler_configuracoes(path, excel_book=None):
    if excel_book is not None:
        workbook = copiar_planilha(excel_book)
    else:
        validar_planilha_fechada(path)
        workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if CONFIG_SHEET not in workbook.sheetnames:
            raise ValueError(f"A aba {CONFIG_SHEET} não existe na planilha.")
        aba = workbook[CONFIG_SHEET]
        if [aba.cell(1, col).value for col in (1, 2)] != CONFIG_HEADERS:
            raise ValueError("Os cabeçalhos da aba Configurações não são os esperados.")
        return validar_configuracoes(aba["B2"].value, aba["B3"].value)
    finally:
        workbook.close()


def validar_configuracoes(intervalo, ativa):
    if isinstance(intervalo, bool) or not isinstance(intervalo, (int, float)) or intervalo not in INTERVALS:
        raise ValueError("O intervalo deve ser 5, 15, 30, 45 ou 60 minutos.")
    if not isinstance(ativa, str) or ativa.strip().casefold() not in ("sim", "não", "nao"):
        raise ValueError("Automação ativa deve ser Sim ou Não.")
    return int(intervalo), ativa.strip().casefold() == "sim"


def ler_configuracoes_atuais(path):
    book = encontrar_planilha(path)
    if book is None:
        return ler_configuracoes(path)
    sheet = book.Worksheets(CONFIG_SHEET)
    if [sheet.Cells(1, col).Value for col in (1, 2)] != CONFIG_HEADERS:
        raise ValueError("Os cabeçalhos da aba Configurações não são os esperados.")
    return validar_configuracoes(sheet.Cells(2, 2).Value, sheet.Cells(3, 2).Value)


def ler_cache(path):
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as file:
        data = json.load(file)
    if not isinstance(data, dict) or data.get("versao") != 1 or not isinstance(data.get("ces"), dict):
        raise ValueError(f"Arquivo de dados inválido: {path}")
    for ce, registro in data["ces"].items():
        if not isinstance(ce, str) or len(ce) != 15 or not ce.isascii() or not ce.isdigit():
            raise ValueError(f"CE inválido no arquivo de dados: {path}")
        if not isinstance(registro, dict) or not isinstance(registro.get("consultado_em"), str):
            raise ValueError(f"Registro inválido no arquivo de dados: {path}")
        status = registro.get("status")
        if not isinstance(status, dict):
            raise ValueError(f"Status inválido no arquivo de dados: {path}")
        mensagens, linhas = status.get("mensagens"), status.get("linhas")
        if not isinstance(mensagens, list) or not all(isinstance(m, str) for m in mensagens):
            raise ValueError(f"Mensagens inválidas no arquivo de dados: {path}")
        if not isinstance(linhas, list) or not all(
            isinstance(linha, list) and len(linha) == 17
            and all(isinstance(value, str) for value in linha)
            and linha[2] == ce
            for linha in linhas
        ):
            raise ValueError(f"Resultados inválidos no arquivo de dados: {path}")
    return data["ces"]


def validar_resultados(aba):
    headers = [aba.cell(1, i).value for i in range(1, 20)]
    if headers not in (RESULT_HEADERS, LEGACY_RESULT_HEADERS):
        raise ValueError("Os cabeçalhos da aba Resultados não são os esperados.")
    if "TabelaResultados" not in aba.tables:
        raise ValueError("A tabela de resultados não foi encontrada.")
    return headers


def linha_resultado(linha, registro, headers):
    valores = dict(zip(PORTAL_HEADERS, linha))
    valores["MENSAGEM DO PORTAL"] = "; ".join(registro["status"]["mensagens"])
    valores["ÚLTIMA CONSULTA"] = registro["consultado_em"]
    return [valores[header] for header in headers]


def atualizar_resultados(aba, atualizacoes):
    headers = validar_resultados(aba)
    ce_index = headers.index("CE")
    existentes = []
    for row in aba.iter_rows(min_row=2, max_col=19, values_only=True):
        if any(value is not None for value in row) and str(row[ce_index]) not in atualizacoes:
            existentes.append(list(row))
    novos = []
    for ce, registro in atualizacoes.items():
        status = registro["status"]
        rows = status["linhas"] or [[None, None, ce] + [None] * 14]
        for row in rows:
            novos.append(linha_resultado(row, registro, headers))
    todos = existentes + novos
    limite = max(aba.max_row, len(todos) + 1, 2)
    for row in range(2, limite + 1):
        values = todos[row - 2] if row - 2 < len(todos) else [None] * 19
        for col, value in enumerate(values, 1):
            cell = aba.cell(row, col)
            cell.value = value
            if row - 2 >= len(existentes) and isinstance(value, str):
                cell.data_type = "s"
    tabela = aba.tables["TabelaResultados"]
    tabela.ref = f"A1:{get_column_letter(19)}{max(2, len(todos) + 1)}"
    if tabela.autoFilter is not None:
        tabela.autoFilter.ref = tabela.ref


def recuperar_resultados(monitor, resultados, selecionados, cache):
    atualizacoes = {}
    headers = validar_resultados(resultados)
    ce_index = headers.index("CE")
    for row, ce in selecionados:
        registro = cache.get(ce)
        if registro is None:
            continue
        anterior = monitor.cell(row, 3).value
        try:
            cache_time = datetime.fromisoformat(registro["consultado_em"])
            sheet_time = datetime.fromisoformat(anterior) if anterior else None
            if cache_time.utcoffset() is None or (sheet_time and sheet_time.utcoffset() is None):
                raise ValueError("Data sem fuso horário.")
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Data de consulta inválida na linha {row} ou no JSON.") from exc
        if sheet_time and sheet_time > cache_time:
            raise ValueError(f"A planilha contém uma consulta mais recente que o JSON na linha {row}.")

        status = registro["status"]
        rows = status["linhas"] or [[None, None, ce] + [None] * 14]
        esperados = [
            [value if value != "" else None for value in linha_resultado(linha, registro, headers)]
            for linha in rows
        ]
        existentes = [
            [value if value != "" else None for value in values]
            for values in resultados.iter_rows(min_row=2, max_col=19, values_only=True)
            if values[ce_index] == ce
        ]
        if anterior != registro["consultado_em"] or existentes != esperados:
            atualizacoes[ce] = registro
            monitor.cell(row, 3).value = registro["consultado_em"]
            monitor.cell(row, 4).value = "Resultado recuperado"
    if atualizacoes:
        atualizar_resultados(resultados, atualizacoes)
    return bool(atualizacoes)


def salvar_json(path, dados):
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as file:
            temporario = Path(file.name)
            json.dump({"versao": 1, "ces": dados}, file, ensure_ascii=False, indent=2)
            file.write("\n")
        os.replace(temporario, path)
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)


def salvar_planilha(workbook, path, original_stat):
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=path.parent, delete=False) as file:
            temporario = Path(file.name)
        workbook.save(temporario)
        stat = path.stat()
        if (stat.st_mtime_ns, stat.st_size) != original_stat:
            raise EdicaoConcorrente("A planilha foi alterada durante a consulta; nenhuma alteração foi gravada nela.")
        os.replace(temporario, path)
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)


def copiar_por_etapa(planilha):
    with planilha_por_ciclo(planilha) as (book, _):
        return copiar_planilha(book)


def aplicar_por_etapa(planilha, antes, depois):
    with planilha_por_ciclo(planilha) as (book, temporaria):
        aplicar_alteracoes(book, antes, depois)
        if temporaria:
            book.Save()
    return temporaria


class ParadaSolicitada(Exception):
    pass


async def consultar_selecionados(selecionados, cache, monitor, visivel, parada, atualizar_status=None):
    sucessos = {}
    erros = 0

    async def consultas():
        nonlocal erros
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(channel="msedge", headless=not visivel)
            try:
                page = await browser.new_page()
                page.set_default_timeout(TIMEOUT)
                for row, ce in selecionados:
                    try:
                        status = await consultar(page, ce)
                        consultado_em = datetime.now().astimezone().isoformat(timespec="seconds")
                        anterior = cache.get(ce)
                        situacao = "Primeira consulta" if anterior is None else (
                            "Sem mudança" if anterior.get("status") == status else "Mudança detectada"
                        )
                        sucessos[ce] = {"consultado_em": consultado_em, "status": status}
                        monitor.cell(row, 3).value = consultado_em
                        monitor.cell(row, 4).value = situacao
                        print(f"CE {ce}: {situacao}.")
                    except (PlaywrightError, ValueError, RuntimeError) as exc:
                        erros += 1
                        monitor.cell(row, 4).value = "Erro na consulta"
                        print(f"CE {ce}: erro na consulta: {exc}", file=sys.stderr)
            finally:
                await browser.close()

    async def aguardar_parada():
        proximo_status = 0.0
        while parada is None or win32event.WaitForSingleObject(parada, 0) != win32event.WAIT_OBJECT_0:
            if atualizar_status is not None and time.monotonic() >= proximo_status:
                atualizar_status()
                proximo_status = time.monotonic() + 5
            await asyncio.sleep(0.2)

    tarefa = asyncio.create_task(consultas())
    if parada is None and atualizar_status is None:
        await tarefa
    else:
        espera = asyncio.create_task(aguardar_parada())
        try:
            concluidas, _ = await asyncio.wait((tarefa, espera), return_when=asyncio.FIRST_COMPLETED)
            if espera in concluidas:
                try:
                    await espera
                finally:
                    tarefa.cancel()
                    try:
                        await tarefa
                    except asyncio.CancelledError:
                        pass
                raise ParadaSolicitada()
            await tarefa
        finally:
            espera.cancel()
            try:
                await espera
            except asyncio.CancelledError:
                pass
    return sucessos, erros


def executar_uma_vez(planilha, dados_path, visivel, excel_book=None, por_etapas=False, parada=None,
                     atualizar_status=None):
    cache = ler_cache(dados_path)
    if excel_book is None and not por_etapas:
        validar_planilha_fechada(planilha)
        stat = planilha.stat()
        original_stat = (stat.st_mtime_ns, stat.st_size)
        workbook = load_workbook(planilha)
        antes = None
    else:
        antes = copiar_por_etapa(planilha) if por_etapas else copiar_planilha(excel_book)
        try:
            snapshot = BytesIO()
            antes.save(snapshot)
            snapshot.seek(0)
            workbook = load_workbook(snapshot)
        except (OSError, TypeError, ValueError):
            antes.close()
            raise
    try:
        if "Monitoramento" not in workbook or "Resultados" not in workbook:
            raise ValueError("A planilha precisa conter as abas Monitoramento e Resultados.")
        monitor = workbook["Monitoramento"]
        resultados = workbook["Resultados"]
        selecionados, selecoes_alteradas = selecionar_ces(monitor)
        validar_resultados(resultados)
        recuperados = recuperar_resultados(monitor, resultados, selecionados, cache)
        if recuperados or selecoes_alteradas:
            if excel_book is not None or por_etapas:
                temporaria = aplicar_por_etapa(planilha, antes, workbook) if por_etapas else False
                if not por_etapas:
                    aplicar_alteracoes(excel_book, antes, workbook)
                print("Avisos e resultados anteriores atualizados." if temporaria else
                      "Avisos e resultados anteriores atualizados no Excel; salve a planilha para mantê-los.")
                if not selecionados:
                    print("Nenhum CE marcado como Sim.")
                    return 0
                antes.close()
                workbook.close()
                antes = copiar_por_etapa(planilha) if por_etapas else copiar_planilha(excel_book)
                snapshot = BytesIO()
                antes.save(snapshot)
                snapshot.seek(0)
                workbook = load_workbook(snapshot)
                monitor = workbook["Monitoramento"]
                resultados = workbook["Resultados"]
                selecionados, selecoes_alteradas = selecionar_ces(monitor)
                validar_resultados(resultados)
            else:
                salvar_planilha(workbook, planilha, original_stat)
                stat = planilha.stat()
                original_stat = (stat.st_mtime_ns, stat.st_size)
                print("Avisos e resultados anteriores atualizados.")
        if not selecionados:
            if selecoes_alteradas and (excel_book is not None or por_etapas):
                if por_etapas:
                    temporaria = aplicar_por_etapa(planilha, antes, workbook)
                else:
                    aplicar_alteracoes(excel_book, antes, workbook)
                    temporaria = False
                print("Avisos atualizados." if temporaria else
                      "Avisos atualizados no Excel; salve a planilha para mantê-los.")
            print("Nenhum CE marcado como Sim.")
            return 0
        sucessos, erros = asyncio.run(
            consultar_selecionados(selecionados, cache, monitor, visivel, parada, atualizar_status)
        )
        atualizar_resultados(resultados, sucessos)
        cache.update(sucessos)
        if sucessos:
            salvar_json(dados_path, cache)
        try:
            if por_etapas:
                temporaria = aplicar_por_etapa(planilha, antes, workbook)
            elif excel_book is not None:
                aplicar_alteracoes(excel_book, antes, workbook)
            else:
                salvar_planilha(workbook, planilha, original_stat)
        except (OSError, RuntimeError, com_error):
            if sucessos:
                print("O JSON foi atualizado, mas a planilha não foi atualizada; execute novamente.", file=sys.stderr)
            raise
        if excel_book is not None or (por_etapas and not temporaria):
            print("Resultados atualizados no Excel; salve a planilha para mantê-los.")
        else:
            print(f"Planilha atualizada: {planilha}")
        return 1 if erros else 0
    finally:
        workbook.close()
        if antes is not None:
            antes.close()


class MonitorEmExecucao(RuntimeError):
    pass


@contextmanager
def bloquear_execucao(path):
    with path.open("a+b") as lock_file:
        lock_file.seek(0)
        try:
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise MonitorEmExecucao("Já existe uma execução do monitor em andamento.") from exc
        try:
            yield
        finally:
            lock_file.seek(0)
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)


def executar_ciclo(planilha, dados_path, visivel, continuo, parada=None, atualizar_status=None):
    por_etapas = planilha.suffix.lower() == ".xlsm"
    contexto = planilha_por_ciclo(planilha) if por_etapas else nullcontext((encontrar_planilha(planilha), False))
    with contexto as (excel_book, _):
        garantir_aba_configuracoes(planilha, excel_book)
        intervalo, ativa = ler_configuracoes(planilha, excel_book)
    if continuo and not ativa:
        print("Automação desativada na aba Configurações; nenhuma consulta foi iniciada.")
        return intervalo, False, 0
    resultado = executar_uma_vez(
        planilha, dados_path, visivel, None if por_etapas else excel_book, por_etapas, parada,
        atualizar_status,
    )
    return intervalo, True, resultado


def informar(texto, oculto):
    print(texto, flush=True)
    if oculto:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, texto, "Monitoramento TCP", 0x40)


class EstadoMonitor:
    def __init__(self, path, planilha):
        self.path = path
        self.planilha = planilha
        self.situacao = "Iniciando"
        self.ultima = "-"
        self.proxima = "-"

    def publicar(self):
        temporario = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-16", dir=self.path.parent,
                                             delete=False, newline="\n") as file:
                temporario = Path(file.name)
                file.write("\n".join((str(self.planilha), self.situacao, self.ultima, self.proxima)) + "\n")
            for tentativa in range(3):
                try:
                    os.replace(temporario, self.path)
                    break
                except PermissionError:
                    if tentativa == 2:
                        raise
                    time.sleep(0.1)
        finally:
            if temporario is not None:
                temporario.unlink(missing_ok=True)


def executar_continuamente(planilha, dados_path, visivel, eventos, estado):
    parada, consulta = eventos
    falhou = False
    aviso_acesso = False
    try:
        while True:
            if win32event.WaitForSingleObject(parada, 0) == win32event.WAIT_OBJECT_0:
                break
            estado.situacao = "Consultando"
            estado.proxima = "-"
            estado.publicar()
            try:
                intervalo, ativa, _ = executar_ciclo(
                    planilha, dados_path, visivel, True, parada, estado.publicar
                )
                if aviso_acesso:
                    print("Acesso ao Excel restabelecido.")
                    aviso_acesso = False
            except ExcelIndisponivel as exc:
                if not aviso_acesso:
                    print(f"Aviso: {exc} Tentando novamente.", file=sys.stderr)
                    aviso_acesso = True
                estado.situacao = "Aguardando Excel"
                estado.publicar()
                if win32event.WaitForSingleObject(parada, 5000) == win32event.WAIT_OBJECT_0:
                    return
                continue
            except EdicaoConcorrente as exc:
                print(f"Aviso: {exc} O monitor tentará novamente no próximo ciclo.", file=sys.stderr)
                aviso_leitura = False
                while True:
                    try:
                        intervalo, ativa = ler_configuracoes_atuais(planilha)
                        break
                    except (com_error, OSError, RuntimeError) as leitura:
                        if not aviso_leitura:
                            print(f"Aviso: aguardando acesso às configurações: {leitura}", file=sys.stderr)
                            aviso_leitura = True
                        estado.situacao = "Aguardando Excel"
                        estado.publicar()
                        if win32event.WaitForSingleObject(parada, 5000) == win32event.WAIT_OBJECT_0:
                            return
            if not ativa:
                print("Automação desativada na aba Configurações; monitoramento encerrado.")
                break
            fim_monotonic = time.monotonic()
            fim = datetime.now().astimezone()
            estado.ultima = fim.strftime("%d/%m/%Y %H:%M:%S")
            aviso_leitura = False
            print(f"Próxima consulta em {intervalo} minutos.", flush=True)
            while True:
                restante = max(0.0, fim_monotonic + intervalo * 60 - time.monotonic())
                proxima = fim.timestamp() + intervalo * 60
                estado.situacao = "Aguardando Excel" if aviso_leitura else "Ativo"
                estado.proxima = "-" if aviso_leitura else datetime.fromtimestamp(proxima).strftime("%d/%m/%Y %H:%M:%S")
                estado.publicar()
                evento = win32event.WaitForMultipleObjects(
                    (parada, consulta), False, int(min(restante, 5) * 1000)
                )
                if evento == win32event.WAIT_OBJECT_0:
                    return
                if evento == win32event.WAIT_OBJECT_0 + 1:
                    print("Consulta imediata solicitada.")
                    break
                try:
                    novo_intervalo, ativa = ler_configuracoes_atuais(planilha)
                except (com_error, OSError, RuntimeError) as exc:
                    if not aviso_leitura:
                        print(f"Aviso: não foi possível ler as configurações; tentando novamente: {exc}",
                              file=sys.stderr)
                        aviso_leitura = True
                    time.sleep(1)
                    continue
                if novo_intervalo != intervalo:
                    intervalo = novo_intervalo
                    print(f"Intervalo alterado para {intervalo} minutos.", flush=True)
                if aviso_leitura:
                    print("Leitura das configurações restabelecida.")
                    aviso_leitura = False
                if not ativa:
                    print("Automação desativada na aba Configurações; monitoramento encerrado.")
                    return
                if time.monotonic() >= fim_monotonic + intervalo * 60:
                    break
    except ParadaSolicitada:
        print("Consulta interrompida a pedido do usuário.")
    except Exception:
        falhou = True
        raise
    finally:
        estado.situacao = "Erro" if falhou else "Parado"
        estado.proxima = "-"
        estado.publicar()


def main():
    parser = argparse.ArgumentParser(description="Consulta os CEs marcados na planilha.")
    parser.add_argument("--planilha", type=Path, default=Path(__file__).parent / "tcp_monitoring_template.xlsx")
    parser.add_argument("--dados", type=Path, default=Path(__file__).parent / "tcp_monitoring_data.json")
    parser.add_argument("--visivel", action="store_true", help="Mostra o Edge durante a consulta.")
    parser.add_argument("--loop", action="store_true", help="Repete consultas conforme a aba Configurações.")
    parser.add_argument("--pausar", action="store_true", help="Mantém a janela aberta após a execução.")
    parser.add_argument("--oculto", action="store_true", help="Registra a execução em arquivo.")
    parser.add_argument("--parar", action="store_true", help="Solicita a parada do monitor contínuo.")
    args = parser.parse_args()
    try:
        planilha = args.planilha.resolve()
        dados_path = args.dados.resolve()
        if planilha == dados_path:
            raise ValueError("A planilha e o arquivo JSON devem ser diferentes.")
        if planilha.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("A planilha deve ter extensão .xlsx ou .xlsm.")
        if args.parar:
            mensagem = ("Solicitação de parada enviada." if enviar(planilha, "parar")
                        else "Nenhum monitoramento compatível está em execução para esta planilha.")
            informar(mensagem, args.oculto)
            return 0
        with bloquear_execucao(Path(__file__).parent / ".tcp_monitor.lock"):
            if args.loop:
                with controles(planilha) as eventos:
                    estado = EstadoMonitor(planilha.parent / "tcp_monitor_status.txt", planilha)
                    executar_continuamente(planilha, dados_path, args.visivel, eventos, estado)
                return 0
            estado = EstadoMonitor(planilha.parent / "tcp_monitor_status.txt", planilha)
            estado.situacao = "Consultando"
            estado.publicar()
            try:
                _, _, resultado = executar_ciclo(
                    planilha, dados_path, args.visivel, False, atualizar_status=estado.publicar
                )
                estado.ultima = datetime.now().astimezone().strftime("%d/%m/%Y %H:%M:%S")
                estado.situacao = "Erro" if resultado else "Parado"
                return resultado
            except Exception:
                estado.situacao = "Erro"
                raise
            finally:
                estado.publicar()
    except MonitorEmExecucao:
        if args.loop:
            informar("Já existe um monitoramento em andamento.", args.oculto)
            return 0
        if enviar(planilha, "consultar"):
            informar("Consulta imediata solicitada ao monitor em andamento.", args.oculto)
            return 0
        informar(
            "Já existe uma execução em andamento, mas ela não aceita solicitações. "
            "Aguarde seu término ou encerre o monitor antigo antes de iniciar a nova versão.",
            args.oculto,
        )
        return 0
    except KeyboardInterrupt:
        print("\nMonitoramento encerrado.")
        return 0
    except (OSError, ValueError, RuntimeError, PlaywrightError, com_error, json.JSONDecodeError) as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    if "--oculto" in sys.argv:
        import ctypes

        log_path = Path(__file__).parent / "tcp_monitor.log"
        try:
            with log_path.open("a", encoding="utf-8", buffering=1) as log:
                with redirect_stdout(log), redirect_stderr(log):
                    print(f"\nInício: {datetime.now().astimezone().isoformat(timespec='seconds')}")
                    try:
                        exit_code = main()
                    except Exception:
                        traceback.print_exc()
                        exit_code = 1
            if exit_code:
                ctypes.windll.user32.MessageBoxW(
                    0, f"O monitor foi encerrado com erro. Consulte o arquivo {log_path}.",
                    "Monitoramento TCP", 0x10,
                )
        except OSError as exc:
            ctypes.windll.user32.MessageBoxW(
                0, f"Não foi possível registrar a execução: {exc}", "Monitoramento TCP", 0x10,
            )
            exit_code = 1
    else:
        exit_code = main()
    if "--pausar" in sys.argv:
        try:
            input("Pressione Enter para fechar esta janela...")
        except EOFError:
            print("Não foi possível aguardar a entrada; a janela será fechada.", file=sys.stderr)
    sys.exit(exit_code)
