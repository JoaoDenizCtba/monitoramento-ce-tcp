import os
import tempfile
from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client
from openpyxl import load_workbook
from pywintypes import com_error


class EdicaoConcorrente(RuntimeError):
    pass


class ExcelIndisponivel(RuntimeError):
    pass


def encontrar_planilha(path):
    rot = pythoncom.GetRunningObjectTable()
    context = pythoncom.CreateBindCtx(0)
    target_path = os.path.normcase(os.path.realpath(path))
    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if not os.path.isabs(name) or os.path.normcase(os.path.realpath(name)) != target_path:
            continue
        try:
            book = win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            nome = book.FullName
            somente_leitura = book.ReadOnly
        except (AttributeError, com_error) as exc:
            raise ExcelIndisponivel("Não foi possível acessar a planilha aberta no Excel.") from exc
        if os.path.normcase(os.path.realpath(nome)) == target_path:
            if somente_leitura:
                raise RuntimeError("A planilha está aberta somente para leitura no Excel.")
            return book
    return None


@contextmanager
def planilha_por_ciclo(path):
    book = encontrar_planilha(path)
    if book is not None:
        yield book, False
        return

    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.AutomationSecurity = 3
        app.EnableEvents = False
        book = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False
        )
        try:
            if book.ReadOnly:
                raise RuntimeError("Não foi possível abrir a planilha para edição no Excel.")
            yield book, True
        finally:
            book.Close(SaveChanges=False)
    finally:
        app.Quit()


def copiar_planilha(book):
    with tempfile.TemporaryDirectory() as directory:
        suffix = Path(book.FullName).suffix.lower()
        path = Path(directory) / f"snapshot{suffix}"
        book.SaveCopyAs(str(path))
        return load_workbook(path)


def assinatura(workbook):
    result = []
    for name in ("Monitoramento", "Resultados", "Configurações"):
        if name not in workbook:
            raise ValueError(f"A aba {name} não existe na planilha.")
        sheet = workbook[name]
        columns = 19 if name == "Resultados" else 4 if name == "Monitoramento" else 2
        rows = sheet.max_row if name != "Configurações" else 3
        cells = tuple(
            (cell.value, cell.data_type)
            for row in sheet.iter_rows(min_row=1, max_row=rows, max_col=columns)
            for cell in row
        )
        tables = tuple((table.name, table.ref) for table in sheet.tables.values())
        result.append((name, rows, cells, tables))
    return tuple(result)


def aplicar_alteracoes(book, antes, depois):
    atual = copiar_planilha(book)
    try:
        if assinatura(atual) != assinatura(antes):
            raise EdicaoConcorrente(
                "A planilha mudou durante a consulta; nenhuma alteração foi aplicada ao Excel. "
                "Revise as alterações e execute novamente."
            )
    finally:
        atual.close()

    try:
        for name in ("Monitoramento", "Resultados"):
            old, new = antes[name], depois[name]
            sheet = book.Worksheets(name)
            columns = 4 if name == "Monitoramento" else 19
            for row in range(2, max(old.max_row, new.max_row) + 1):
                for col in range(1, columns + 1):
                    previous, cell = old.cell(row, col), new.cell(row, col)
                    if (previous.value, previous.data_type) == (cell.value, cell.data_type):
                        continue
                    target = sheet.Cells(row, col)
                    if cell.data_type == "s":
                        target.NumberFormat = "@"
                        target.Value = cell.value
                    elif cell.data_type == "f":
                        target.Formula = cell.value
                    else:
                        target.Value = cell.value
            if name == "Resultados":
                ref = new.tables["TabelaResultados"].ref
                sheet.ListObjects("TabelaResultados").Resize(sheet.Range(ref))
    except com_error as exc:
        raise RuntimeError(
            "Não foi possível aplicar todas as alterações no Excel; a planilha pode ter "
            "alterações parciais não salvas. Revise a planilha e execute novamente."
        ) from exc
