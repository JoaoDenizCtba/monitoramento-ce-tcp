import asyncio
import subprocess
import sys
import tempfile
import unittest
from contextlib import nullcontext
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import win32event

import monitor_control
import tcp_monitor as monitor
from test_tcp_monitor import CE, planilha, registro


class EventTests(unittest.TestCase):
    def test_sinais_sao_entregues_apenas_enquanto_monitor_esta_ativo(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsm"
            self.assertFalse(monitor_control.enviar(path, "parar"))
            with monitor_control.controles(path) as (parar, consultar):
                self.assertTrue(monitor_control.enviar(path, "consultar"))
                self.assertEqual(win32event.WaitForSingleObject(consultar, 0), win32event.WAIT_OBJECT_0)
                self.assertEqual(win32event.WaitForSingleObject(parar, 0), win32event.WAIT_TIMEOUT)
                self.assertTrue(monitor_control.enviar(path, "parar"))
                self.assertEqual(win32event.WaitForSingleObject(parar, 0), win32event.WAIT_OBJECT_0)
            self.assertFalse(monitor_control.enviar(path, "consultar"))

    def test_comando_pode_ser_enviado_por_outro_processo(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsm"
            with monitor_control.controles(path) as (parar, _):
                result = subprocess.run(
                    [sys.executable, "-c",
                     "import monitor_control, sys; sys.exit(0 if monitor_control.enviar(sys.argv[1], 'parar') else 1)",
                     str(path)], capture_output=True, text=True, timeout=10,
                )
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(win32event.WaitForSingleObject(parar, 0), win32event.WAIT_OBJECT_0)

    def test_status_usa_unicode_e_nao_altera_outros_arquivos(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "status.txt"
            planilha_path = Path(directory) / "planilha.xlsm"
            estado = monitor.EstadoMonitor(path, planilha_path)
            estado.situacao = "Ativo"
            estado.publicar()
            self.assertEqual(path.read_text(encoding="utf-16").splitlines(),
                             [str(planilha_path), "Ativo", "-", "-"])

    def test_consulta_avulsa_atualiza_status_apos_erro_anterior(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsm"
            args = ["tcp_monitor.py", "--planilha", str(path)]
            with patch.object(sys, "argv", args), \
                    patch.object(monitor, "bloquear_execucao", return_value=nullcontext()), \
                    patch.object(monitor, "executar_ciclo", return_value=(5, True, 0)):
                self.assertEqual(monitor.main(), 0)
            status = (Path(directory) / "tcp_monitor_status.txt").read_text(encoding="utf-16").splitlines()
            self.assertEqual(status[1], "Parado")
            self.assertNotEqual(status[2], "-")

    def test_alteracao_de_intervalo_reprograma_a_partir_do_fim(self):
        estado = MagicMock()
        momentos = []
        estado.publicar.side_effect = lambda: momentos.append((estado.situacao, estado.proxima))
        with patch.object(monitor, "executar_ciclo", return_value=(5, True, 0)), \
                patch.object(monitor, "ler_configuracoes_atuais", return_value=(30, True)), \
                patch.object(monitor.win32event, "WaitForSingleObject", return_value=win32event.WAIT_TIMEOUT), \
                patch.object(monitor.win32event, "WaitForMultipleObjects",
                             side_effect=[win32event.WAIT_TIMEOUT, win32event.WAIT_OBJECT_0]):
            monitor.executar_continuamente(Path("teste.xlsm"), Path("dados.json"), False, (1, 2), estado)
        ativos = [value for state, value in momentos if state == "Ativo"]
        primeira, segunda = [monitor.datetime.strptime(value, "%d/%m/%Y %H:%M:%S") for value in ativos]
        self.assertEqual((segunda - primeira).total_seconds(), 25 * 60)
        self.assertEqual(estado.situacao, "Parado")

    def test_edicao_concorrente_nao_encerra_o_monitor(self):
        estado = MagicMock()
        with patch.object(monitor, "executar_ciclo", side_effect=monitor.EdicaoConcorrente("mudou")), \
                patch.object(monitor, "ler_configuracoes_atuais", return_value=(5, True)), \
                patch.object(monitor.win32event, "WaitForSingleObject", return_value=win32event.WAIT_TIMEOUT), \
                patch.object(monitor.win32event, "WaitForMultipleObjects", return_value=win32event.WAIT_OBJECT_0):
            monitor.executar_continuamente(Path("teste.xlsm"), Path("dados.json"), False, (1, 2), estado)
        self.assertEqual(estado.situacao, "Parado")
        self.assertGreaterEqual(estado.publicar.call_count, 3)

    def test_acesso_transitorio_ao_excel_e_tentado_novamente(self):
        estado = MagicMock()
        with patch.object(monitor, "executar_ciclo",
                          side_effect=[monitor.ExcelIndisponivel("indisponível"), (5, True, 0)]) as ciclo, \
                patch.object(monitor.win32event, "WaitForSingleObject", return_value=win32event.WAIT_TIMEOUT), \
                patch.object(monitor.win32event, "WaitForMultipleObjects", return_value=win32event.WAIT_OBJECT_0):
            monitor.executar_continuamente(Path("teste.xlsm"), Path("dados.json"), False, (1, 2), estado)
        self.assertEqual(ciclo.call_count, 2)
        self.assertEqual(estado.situacao, "Parado")


class CancellationTests(unittest.IsolatedAsyncioTestCase):
    async def test_parada_cancela_consulta_e_fecha_navegador(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "teste.xlsm"
            page = MagicMock()
            browser = MagicMock()
            browser.new_page = AsyncMock(return_value=page)
            browser.close = AsyncMock()
            playwright = MagicMock()
            playwright.chromium.launch = AsyncMock(return_value=browser)
            context = MagicMock()
            context.__aenter__ = AsyncMock(return_value=playwright)
            context.__aexit__ = AsyncMock(return_value=False)
            iniciada = asyncio.Event()

            async def consulta_lenta(_page, _ce):
                iniciada.set()
                await asyncio.Event().wait()

            with monitor_control.controles(path) as (parar, _), \
                    patch.object(monitor, "async_playwright", return_value=context), \
                    patch.object(monitor, "consultar", side_effect=consulta_lenta):
                tarefa = asyncio.create_task(monitor.consultar_selecionados(
                    [(2, CE)], {}, planilha()["Monitoramento"], False, parar
                ))
                await asyncio.wait_for(iniciada.wait(), 2)
                self.assertTrue(monitor_control.enviar(path, "parar"))
                with self.assertRaises(monitor.ParadaSolicitada):
                    await asyncio.wait_for(tarefa, 2)
            browser.close.assert_awaited_once()

    async def test_consulta_concluida_e_retornada(self):
        page = MagicMock()
        browser = MagicMock()
        browser.new_page = AsyncMock(return_value=page)
        browser.close = AsyncMock()
        playwright = MagicMock()
        playwright.chromium.launch = AsyncMock(return_value=browser)
        context = MagicMock()
        context.__aenter__ = AsyncMock(return_value=playwright)
        context.__aexit__ = AsyncMock(return_value=False)
        sheet = planilha()["Monitoramento"]
        with patch.object(monitor, "async_playwright", return_value=context), \
                patch.object(monitor, "consultar", new_callable=AsyncMock, return_value=registro()["status"]):
            sucessos, erros = await monitor.consultar_selecionados([(2, CE)], {}, sheet, False, None)
        self.assertEqual(sucessos[CE]["status"], registro()["status"])
        self.assertEqual(erros, 0)
        self.assertEqual(sheet["D2"].value, "Primeira consulta")
        browser.close.assert_awaited_once()


if __name__ == "__main__":
    unittest.main()
