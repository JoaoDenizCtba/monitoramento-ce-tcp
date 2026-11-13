import hashlib
import os
from contextlib import contextmanager

import pywintypes
import win32api
import win32con
import win32event


def nome_evento(planilha, comando):
    caminho = os.path.normcase(os.path.realpath(planilha))
    digest = hashlib.sha256(os.fsencode(caminho)).hexdigest()
    return f"Local\\MonitorTCP-{digest}-{comando}"


@contextmanager
def controles(planilha):
    parar = win32event.CreateEvent(None, False, False, nome_evento(planilha, "parar"))
    try:
        consultar = win32event.CreateEvent(None, False, False, nome_evento(planilha, "consultar"))
        try:
            yield parar, consultar
        finally:
            win32api.CloseHandle(consultar)
    finally:
        win32api.CloseHandle(parar)


def enviar(planilha, comando):
    try:
        evento = win32event.OpenEvent(win32con.EVENT_MODIFY_STATE, False, nome_evento(planilha, comando))
    except pywintypes.error as exc:
        if exc.winerror == 2:
            return False
        raise
    try:
        win32event.SetEvent(evento)
    finally:
        win32api.CloseHandle(evento)
    return True
