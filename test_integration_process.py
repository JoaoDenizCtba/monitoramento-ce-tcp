import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path


ROOT = Path(__file__).parent

RUNNER = r'''
import json
import os
import sys
import time
from pathlib import Path
import win32event
import tcp_monitor as monitor

state_path = Path(os.environ["TEST_STATE"])
mode = os.environ.get("TEST_MODE", "quick")

def save(data):
    temporary = state_path.with_suffix(".tmp")
    temporary.write_text(json.dumps(data), encoding="utf-8")
    temporary.replace(state_path)

def ciclo(planilha, dados, visivel, continuo, parada=None, atualizar_status=None):
    data = json.loads(state_path.read_text(encoding="utf-8"))
    data["count"] = data.get("count", 0) + 1
    data["active"] = data.get("active", 0) + 1
    data["max_active"] = max(data.get("max_active", 0), data["active"])
    save(data)
    if mode == "stop":
        while win32event.WaitForSingleObject(parada, 50) != win32event.WAIT_OBJECT_0:
            time.sleep(0.01)
        raise monitor.ParadaSolicitada()
    time.sleep(0.08)
    data = json.loads(state_path.read_text(encoding="utf-8"))
    data["active"] -= 1
    save(data)
    return (1, True, 0)

def configuracoes(planilha):
    return (1, True)

monitor.executar_ciclo = ciclo
monitor.ler_configuracoes_atuais = configuracoes
sys.exit(monitor.main())
'''


class IsolatedProcessIntegrationTests(unittest.TestCase):
    def isolated(self, mode="quick"):
        temporary = tempfile.TemporaryDirectory(prefix="tcp-monitor-integration-")
        directory = Path(temporary.name)
        for name in ("tcp_monitor.py", "excel_open.py", "monitor_control.py"):
            shutil.copy2(ROOT / name, directory / name)
        (directory / "runner.py").write_text(RUNNER, encoding="utf-8")
        planilha = directory / "fake-planilha.xlsx"
        dados = directory / "fake-dados.json"
        state = directory / "state.json"
        state.write_text("{}", encoding="utf-8")
        env = os.environ.copy()
        env.update(TEST_STATE=str(state), TEST_MODE=mode, PYTHONPATH=str(directory), PYTHONIOENCODING="utf-8")
        base = [sys.executable, str(directory / "runner.py"), "--planilha", str(planilha), "--dados", str(dados)]
        return temporary, state, env, base

    @staticmethod
    def wait_for(predicate, timeout=3):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if predicate():
                return
            time.sleep(0.02)
        raise AssertionError("tempo excedido aguardando condição")

    @staticmethod
    def read_state(path):
        for _ in range(20):
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except (PermissionError, json.JSONDecodeError):
                time.sleep(0.005)
        return json.loads(path.read_text(encoding="utf-8"))

    def finish(self, process, timeout=3):
        out, err = process.communicate(timeout=timeout)
        return process.returncode, out, err

    def test_loop_duplicado_informa_e_nao_cria_outro_loop(self):
        temporary, state, env, base = self.isolated()
        directory = Path(temporary.name)
        first = duplicate = None
        try:
            first = subprocess.Popen(base + ["--loop"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     text=True, encoding="utf-8", cwd=directory, env=env)
            self.wait_for(lambda: self.read_state(state).get("count") == 1)
            duplicate = subprocess.run(base + ["--loop"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       text=True, encoding="utf-8", cwd=directory, env=env, timeout=3)
            self.assertEqual(duplicate.returncode, 0, duplicate.stderr)
            self.assertIn("Já existe um monitoramento em andamento.", duplicate.stdout)
            time.sleep(0.15)
            self.assertEqual(self.read_state(state).get("count"), 1)
        finally:
            if first and first.poll() is None:
                subprocess.run(base + ["--parar"], cwd=directory, env=env, capture_output=True,
                               text=True, encoding="utf-8", timeout=3, check=True)
            if first:
                first.communicate(timeout=3)
            temporary.cleanup()

    def test_consulta_imediata_inicia_novo_ciclo_sem_paralelismo(self):
        temporary, state, env, base = self.isolated()
        directory = Path(temporary.name)
        first = None
        try:
            first = subprocess.Popen(base + ["--loop"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     text=True, encoding="utf-8", cwd=directory, env=env)
            self.wait_for(lambda: self.read_state(state).get("count") == 1)
            request = subprocess.run(base, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     text=True, encoding="utf-8", cwd=directory, env=env, timeout=3)
            self.assertEqual(request.returncode, 0, request.stderr)
            self.assertIn("Consulta imediata solicitada ao monitor em andamento.", request.stdout)
            self.wait_for(lambda: self.read_state(state).get("count") == 2)
            self.assertEqual(self.read_state(state).get("max_active"), 1)
        finally:
            if first and first.poll() is None:
                subprocess.run(base + ["--parar"], cwd=directory, env=env, capture_output=True,
                               text=True, encoding="utf-8", timeout=3, check=True)
            if first:
                first.communicate(timeout=3)
            temporary.cleanup()

    def test_parar_interrompe_consulta_e_publica_parado_em_utf16(self):
        temporary, state, env, base = self.isolated(mode="stop")
        directory = Path(temporary.name)
        first = None
        try:
            first = subprocess.Popen(base + ["--loop"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     text=True, encoding="utf-8", cwd=directory, env=env)
            status = directory / "tcp_monitor_status.txt"
            self.wait_for(lambda: self.read_state(state).get("count") == 1)
            self.wait_for(status.exists)
            stop = subprocess.run(base + ["--parar"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                  text=True, encoding="utf-8", cwd=directory, env=env, timeout=3)
            self.assertEqual(stop.returncode, 0, stop.stderr)
            self.assertIn("Solicitação de parada enviada.", stop.stdout)
            code, out, err = self.finish(first)
            self.assertEqual(code, 0, err)
            self.wait_for(lambda: status.read_text(encoding="utf-16").splitlines()[1] == "Parado")
            self.assertEqual(status.read_text(encoding="utf-16").splitlines()[1], "Parado")
        finally:
            if first and first.poll() is None:
                subprocess.run(base + ["--parar"], cwd=directory, env=env, capture_output=True,
                               text=True, encoding="utf-8", timeout=3, check=True)
                first.communicate(timeout=3)
            temporary.cleanup()


if __name__ == "__main__":
    unittest.main(verbosity=2)
