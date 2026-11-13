import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

import tcp_monitor as monitor


CE = "123456789012345"
OTHER_CE = "123456789012346"
TIME = "2026-11-12T01:00:00-03:00"


def planilha():
    book = Workbook()
    sheet = book.active
    sheet.title = "Monitoramento"
    sheet.append(monitor.MONITOR_HEADERS)
    sheet.append([CE, "Sim", None, None])
    sheet.append([OTHER_CE, "Não", None, None])
    validation = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True)
    sheet.add_data_validation(validation)
    validation.add("B2:B1048576")
    result = book.create_sheet("Resultados")
    result.append(monitor.RESULT_HEADERS)
    result.add_table(Table(displayName="TabelaResultados", ref="A1:S2"))
    monitor.criar_aba_configuracoes(book)
    return book


def registro(ce=CE):
    row = [""] * 17
    row[2] = ce
    return {"consultado_em": TIME, "status": {"mensagens": ["Mensagem"], "linhas": [row]}}


class RecoveryTests(unittest.TestCase):
    def test_recupera_antes_de_tentar_abrir_navegador(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsx"
            cache = Path(directory) / "dados.json"
            book = planilha()
            book.save(path)
            book.close()
            monitor.salvar_json(cache, {CE: registro(), OTHER_CE: registro(OTHER_CE)})
            with patch.object(monitor, "async_playwright", side_effect=RuntimeError("navegador indisponível")):
                with self.assertRaisesRegex(RuntimeError, "navegador indisponível"):
                    monitor.executar_uma_vez(path, cache, False)
            saved = load_workbook(path)
            self.assertEqual(saved["Monitoramento"]["C2"].value, TIME)
            self.assertEqual(saved["Monitoramento"]["D2"].value, "Resultado recuperado")
            self.assertEqual(saved["Resultados"]["E2"].value, CE)
            self.assertNotEqual(saved["Resultados"]["E3"].value, OTHER_CE)
            saved.close()

    def test_nao_sobrescreve_consulta_mais_nova_na_planilha(self):
        book = planilha()
        book["Monitoramento"]["C2"] = "2026-11-12T02:00:00-03:00"
        with self.assertRaisesRegex(ValueError, "mais recente"):
            monitor.recuperar_resultados(
                book["Monitoramento"], book["Resultados"], [(2, CE)], {CE: registro()}
            )
        self.assertIsNone(book["Resultados"]["E2"].value)

    def test_atualiza_preservando_ce_na_ordem_antiga_e_nova(self):
        for headers in (monitor.RESULT_HEADERS, monitor.LEGACY_RESULT_HEADERS):
            with self.subTest(headers=headers):
                book = planilha()
                sheet = book["Resultados"]
                for col, header in enumerate(headers, 1):
                    sheet.cell(1, col).value = header
                monitor.atualizar_resultados(sheet, {CE: registro()})
                row = [sheet.cell(2, col).value for col in range(1, 20)]
                self.assertEqual(row[headers.index("CE")], CE)
                self.assertEqual(row[headers.index("ÚLTIMA CONSULTA")], TIME)
                self.assertEqual(row[headers.index("MENSAGEM DO PORTAL")], "Mensagem")
                book["Monitoramento"].cell(2, 3).value = TIME
                self.assertFalse(monitor.recuperar_resultados(
                    book["Monitoramento"], sheet, [(2, CE)], {CE: registro()}
                ))

    def test_nao_repete_recuperacao_quando_dados_sao_iguais(self):
        book = planilha()
        sheet = book["Monitoramento"]
        results = book["Resultados"]
        self.assertTrue(monitor.recuperar_resultados(sheet, results, [(2, CE)], {CE: registro()}))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsx"
            book.save(path)
            saved = load_workbook(path)
            self.assertFalse(monitor.recuperar_resultados(
                saved["Monitoramento"], saved["Resultados"], [(2, CE)], {CE: registro()}
            ))
            saved.close()

    def test_flag_vazia_sinaliza_pendencia_e_ignora_ce(self):
        book = planilha()
        book["Monitoramento"]["B2"] = None
        selecionados, alterou = monitor.selecionar_ces(book["Monitoramento"])
        self.assertEqual(selecionados, [])
        self.assertTrue(alterou)
        self.assertEqual(book["Monitoramento"]["D2"].value, monitor.SELECAO_PENDENTE)
        book["Monitoramento"]["B2"] = "Não"
        self.assertEqual(monitor.selecionar_ces(book["Monitoramento"]), ([], True))
        self.assertIsNone(book["Monitoramento"]["D2"].value)

    def test_seleciona_ces_com_tabela_deslocada_por_botoes(self):
        sheet = planilha()["Monitoramento"]
        sheet.insert_rows(1, 4)
        sheet.add_table(Table(displayName="TabelaMonitoramento", ref="A5:D7"))
        self.assertEqual(monitor.selecionar_ces(sheet), ([(6, CE)], False))
        sheet["B6"] = None
        self.assertEqual(monitor.selecionar_ces(sheet), ([], True))
        self.assertEqual(sheet["D6"].value, monitor.SELECAO_PENDENTE)

    def test_consulta_selecionados_e_salva_aviso_para_pendentes(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsx"
            cache = Path(directory) / "dados.json"
            book = planilha()
            book["Monitoramento"]["B3"] = None
            book.save(path)
            book.close()

            async def consulta(selecionados, _cache, sheet, _visivel, _parada, _status):
                self.assertEqual(selecionados, [(2, CE)])
                sheet["C2"] = TIME
                sheet["D2"] = "Primeira consulta"
                return {CE: registro()}, 0

            with patch.object(monitor, "consultar_selecionados", side_effect=consulta):
                self.assertEqual(monitor.executar_uma_vez(path, cache, False), 0)
            saved = load_workbook(path)
            self.assertEqual(saved["Monitoramento"]["D3"].value, monitor.SELECAO_PENDENTE)
            self.assertEqual(saved["Resultados"]["E2"].value, CE)
            saved.close()

    def test_sem_ces_selecionados_salva_aviso_sem_abrir_navegador(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsx"
            book = planilha()
            book["Monitoramento"]["B2"] = None
            book.save(path)
            book.close()
            with patch.object(monitor, "consultar_selecionados") as consultar:
                self.assertEqual(monitor.executar_uma_vez(path, Path(directory) / "dados.json", False), 0)
                consultar.assert_not_called()
            saved = load_workbook(path)
            self.assertEqual(saved["Monitoramento"]["D2"].value, monitor.SELECAO_PENDENTE)
            saved.close()

    def test_recusa_tabela_em_colunas_diferentes(self):
        sheet = planilha()["Monitoramento"]
        sheet.add_table(Table(displayName="TabelaMonitoramento", ref="B1:E3"))
        with self.assertRaisesRegex(ValueError, "colunas A até D"):
            monitor.selecionar_ces(sheet)

    def test_prepara_configuracao_5_minutos_sem_proteger_entrada(self):
        book = planilha()
        sheet = book["Configurações"]
        sheet.protection.sheet = False
        sheet["B2"] = 15
        sheet.data_validations.dataValidation[0].formula1 = '"15,30,45,60"'
        monitor.preparar_configuracoes(book)
        self.assertEqual(sheet["B2"].value, 5)
        self.assertTrue(sheet.protection.sheet)
        self.assertTrue(sheet["A2"].protection.locked)
        self.assertFalse(sheet["B2"].protection.locked)
        self.assertFalse(sheet["B3"].protection.locked)
        self.assertEqual(sheet.data_validations.dataValidation[0].formula1, '"5,15,30,45,60"')


class ExcelCycleTests(unittest.TestCase):
    def test_pendencia_sem_ce_selecionado_e_aplicada_no_excel(self):
        anterior = planilha()
        anterior["Monitoramento"]["B2"] = None
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(monitor, "copiar_por_etapa", return_value=anterior), \
                patch.object(monitor, "aplicar_por_etapa", return_value=False) as aplicar, \
                patch.object(monitor, "consultar_selecionados") as consultar:
            path = Path(directory) / "teste.xlsm"
            self.assertEqual(monitor.executar_uma_vez(
                path, Path(directory) / "dados.json", False, por_etapas=True
            ), 0)
            self.assertEqual(aplicar.call_count, 1)
            self.assertEqual(aplicar.call_args.args[2]["Monitoramento"]["D2"].value,
                             monitor.SELECAO_PENDENTE)
            consultar.assert_not_called()

    def test_fecha_instancia_temporaria_antes_de_consultar(self):
        state = {"aberta": False}
        book = MagicMock()

        @contextmanager
        def contexto(_path):
            state["aberta"] = True
            try:
                yield book, True
            finally:
                state["aberta"] = False

        def consultar(*args):
            self.assertFalse(state["aberta"])
            self.assertTrue(args[4])
            return 0

        with patch.object(monitor, "planilha_por_ciclo", side_effect=contexto), \
                patch.object(monitor, "garantir_aba_configuracoes"), \
                patch.object(monitor, "ler_configuracoes", return_value=(5, True)), \
                patch.object(monitor, "executar_uma_vez", side_effect=consultar):
            self.assertEqual(monitor.executar_ciclo(Path("teste.xlsm"), Path("dados.json"), False, True),
                             (5, True, 0))

    def test_aplica_e_salva_apenas_instancia_temporaria(self):
        for temporaria in (False, True):
            with self.subTest(temporaria=temporaria):
                book = MagicMock()

                @contextmanager
                def contexto(_path):
                    yield book, temporaria

                with patch.object(monitor, "planilha_por_ciclo", side_effect=contexto), \
                        patch.object(monitor, "aplicar_alteracoes") as aplicar:
                    self.assertEqual(monitor.aplicar_por_etapa(Path("teste.xlsm"), "antes", "depois"),
                                     temporaria)
                aplicar.assert_called_once_with(book, "antes", "depois")
                self.assertEqual(book.Save.call_count, int(temporaria))

    def test_nao_salva_se_alteracoes_concorrentes_forem_detectadas(self):
        book = MagicMock()

        @contextmanager
        def contexto(_path):
            yield book, True

        with patch.object(monitor, "planilha_por_ciclo", side_effect=contexto), \
                patch.object(monitor, "aplicar_alteracoes", side_effect=RuntimeError("mudou")):
            with self.assertRaisesRegex(RuntimeError, "mudou"):
                monitor.aplicar_por_etapa(Path("teste.xlsm"), "antes", "depois")
        book.Save.assert_not_called()


if __name__ == "__main__":
    unittest.main()
