import unittest
from pathlib import Path
from unittest.mock import MagicMock, call, patch

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

import excel_open


def criar_planilha():
    workbook = Workbook()
    monitor = workbook.active
    monitor.title = "Monitoramento"
    monitor.append(["CE", "Monitorar", "Última consulta", "Situação da consulta"])
    monitor.append(["123456789012345", "Sim", None, None])
    results = workbook.create_sheet("Resultados")
    results.append([f"Coluna {i}" for i in range(1, 20)])
    results.add_table(Table(displayName="TabelaResultados", ref="A1:S2"))
    config = workbook.create_sheet("Configurações")
    config.append(["Configuração", "Valor"])
    config.append(["Intervalo (minutos)", 15])
    config.append(["Automação ativa", "Sim"])
    return workbook


class ExcelOpenTests(unittest.TestCase):
    def test_objeto_do_excel_indisponivel_nao_e_tratado_como_planilha_fechada(self):
        path = Path("teste.xlsm").resolve()
        moniker = MagicMock()
        moniker.GetDisplayName.return_value = str(path)
        rot = MagicMock()
        rot.__iter__.return_value = iter([moniker])
        book = MagicMock(spec=[])
        with patch.object(excel_open.pythoncom, "GetRunningObjectTable", return_value=rot), \
                patch.object(excel_open.pythoncom, "CreateBindCtx"), \
                patch.object(excel_open.win32com.client, "Dispatch", return_value=book):
            with self.assertRaisesRegex(excel_open.ExcelIndisponivel, "acessar a planilha"):
                excel_open.encontrar_planilha(path)

    def test_planilha_aberta_nao_e_salva_nem_fechada_pelo_contexto(self):
        book = MagicMock()
        with patch.object(excel_open, "encontrar_planilha", return_value=book):
            with excel_open.planilha_por_ciclo(Path("teste.xlsm")) as (obtida, temporaria):
                self.assertIs(obtida, book)
                self.assertFalse(temporaria)
        book.Close.assert_not_called()
        book.Save.assert_not_called()

    def test_planilha_fechada_usa_excel_temporario_e_fecha_em_erro(self):
        app = MagicMock()
        book = app.Workbooks.Open.return_value
        book.ReadOnly = False
        with patch.object(excel_open, "encontrar_planilha", return_value=None), \
                patch.object(excel_open.win32com.client, "DispatchEx", return_value=app):
            with self.assertRaisesRegex(RuntimeError, "erro de teste"):
                with excel_open.planilha_por_ciclo(Path("teste.xlsm")) as (obtida, temporaria):
                    self.assertIs(obtida, book)
                    self.assertTrue(temporaria)
                    raise RuntimeError("erro de teste")
        book.Close.assert_called_once_with(SaveChanges=False)
        app.Quit.assert_called_once()

    def test_recusa_alteracao_concorrente_sem_escrever(self):
        anterior = criar_planilha()
        novo = criar_planilha()
        atual = criar_planilha()
        atual["Monitoramento"]["B2"] = "Não"
        book = MagicMock()
        with patch.object(excel_open, "copiar_planilha", return_value=atual):
            with self.assertRaisesRegex(RuntimeError, "mudou durante a consulta"):
                excel_open.aplicar_alteracoes(book, anterior, novo)
        book.Worksheets.assert_not_called()

    def test_texto_do_portal_nao_se_torna_formula(self):
        anterior = criar_planilha()
        novo = criar_planilha()
        atual = criar_planilha()
        novo["Resultados"]["A2"] = "=1+1"
        novo["Resultados"]["A2"].data_type = "s"
        book = MagicMock()
        with patch.object(excel_open, "copiar_planilha", return_value=atual):
            excel_open.aplicar_alteracoes(book, anterior, novo)
        target = book.Worksheets("Resultados").Cells(2, 1)
        self.assertEqual(target.NumberFormat, "@")
        self.assertEqual(target.Value, "=1+1")
        target.Formula.assert_not_called()

    def test_atualiza_linha_deslocada_sem_alterar_o_cabecalho(self):
        anterior = criar_planilha()
        novo = criar_planilha()
        atual = criar_planilha()
        for workbook in (anterior, novo, atual):
            workbook["Monitoramento"].insert_rows(1, 4)
        novo["Monitoramento"]["C6"] = "consulta"
        book = MagicMock()
        with patch.object(excel_open, "copiar_planilha", return_value=atual):
            excel_open.aplicar_alteracoes(book, anterior, novo)
        sheet = book.Worksheets("Monitoramento")
        self.assertEqual(sheet.Cells(6, 3).Value, "consulta")
        sheet.Cells.assert_any_call(6, 3)
        self.assertNotIn(call(5, 3), sheet.Cells.call_args_list)


if __name__ == "__main__":
    unittest.main()
