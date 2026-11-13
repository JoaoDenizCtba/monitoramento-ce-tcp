# Monitoramento de CE - TCP

[Português](#português) | [Español](#español) | [English](#english)

## Português

Automação para consultar CEs selecionados no portal TCP e acompanhar os resultados no Excel. O monitor pode executar uma consulta avulsa ou repetir consultas em intervalos configurados.

### Componentes

- `tcp_monitor.py`: consulta, comparação de resultados e coordenação.
- `excel_open.py`: integração com o Excel e proteção contra alterações concorrentes.
- `monitor_control.py`: sinais de consulta imediata e parada.
- `excel_buttons.bas` e `excel_workbook_events.txt`: código VBA para os botões e o painel.
- `modelo/tcp_monitoring_template.xlsx`: modelo vazio; os rótulos e mensagens da solução estão em português.
- `test_*.py`: testes locais; `requirements.txt`: dependências Python.

### Requisitos

- Windows, Microsoft Excel desktop, Microsoft Edge, acesso ao portal TCP e Python 3.11.
- Dependências instaladas no Python 3.11:

```powershell
py -3.11 -m pip install -r requirements.txt
```

### Preparação

1. Copie `modelo/tcp_monitoring_template.xlsx` para a pasta dos programas, mantendo o nome `tcp_monitoring_template.xlsx`. Use essa cópia para os seus dados.
2. Na aba `Monitoramento`, insira cada CE como **texto com 15 dígitos** e selecione `Sim` ou `Não` em `Monitorar`.
3. Na aba `Configurações`, escolha o intervalo de 5, 15, 30, 45 ou 60 minutos. Para o modo contínuo, escolha `Sim` em `Automação ativa`.

O modelo versionado não deve receber dados reais. O `.gitignore` protege as planilhas de trabalho, mas não impede a publicação acidental se o modelo versionado for editado ou arquivos forem adicionados à força.

### Consulta pelo terminal

Na pasta dos programas:

```powershell
py -3.11 tcp_monitor.py
```

Para repetir consultas enquanto o terminal estiver aberto:

```powershell
py -3.11 tcp_monitor.py --loop
```

O modo contínuo exige `Automação ativa` em `Sim`; para encerrá-lo no terminal, use `Ctrl+C`. Para especificar outra planilha ou outro arquivo de dados, use `--planilha` e `--dados`.

### Botões no Excel

O repositório fornece o **código das macros**, não uma planilha `.xlsm` pronta. Na cópia de trabalho, use **Salvar como** e escolha uma pasta de trabalho habilitada para macro, por exemplo `tcp_monitoring.xlsm`, na mesma pasta dos programas.

No Excel, confirme que as políticas da organização permitem macros. Importe `excel_buttons.bas` pelo editor VBA (`Alt+F11`, `Arquivo > Importar Arquivo`), crie dois botões na aba `Monitoramento` e atribua `ConsultarAgora` e `IniciarMonitoramento`. Execute `PrepararPainel` uma vez para criar o botão de parada e o painel; as células `E1:L4` devem estar vazias, sem formas sobre as áreas de destino. Adicione os procedimentos de `excel_workbook_events.txt` em `ThisWorkbook`, preservando eventos existentes, e salve a planilha.

- **Consultar agora:** executa uma rodada ou solicita uma rodada imediata ao monitor ativo. Se uma consulta estiver em andamento, aguarda seu término.
- **Iniciar monitoramento:** inicia o modo contínuo; apenas uma execução pode ocorrer por vez.
- **Parar monitoramento:** após confirmação, solicita o cancelamento da consulta atual; os resultados dessa rodada são descartados.

O painel é atualizado enquanto o Excel está aberto. `Sem resposta` indica que o status deixou de ser renovado; não confirma que o processo esteja funcionando. A alteração dos arquivos VBA externos não atualiza automaticamente o código já importado.

### Funcionamento e limites

- Apenas CEs com `Sim` são consultados. Um CE sem seleção fica pendente, sem impedir os demais.
- Com a planilha aberta, salve manualmente os resultados e configurações que deseja preservar. Com ela fechada, o monitor usa o Excel para abrir e salvar o arquivo temporariamente.
- O monitor pode continuar com o Excel fechado, desde que o Windows e o processo permaneçam ativos. Não inicia automaticamente após desligamento, reinício ou encerramento da sessão.
- Alterações concorrentes podem impedir uma atualização para proteger os dados. O JSON pode já ter sido atualizado e permitir recuperação na próxima rodada.
- `tcp_monitoring_data.json` guarda resultados para comparação e recuperação; `tcp_monitor.log` registra execuções ocultas; `tcp_monitor_status.txt` alimenta o painel. Preserve esses arquivos conforme sua necessidade.
- Mudanças no portal, falta de acesso, bloqueios do Excel ou políticas corporativas podem impedir a execução. Não contorne políticas de segurança; consulte a TI.

### Privacidade e testes

Planilhas, JSON e logs podem conter informações comerciais. Revise dados e metadados antes de compartilhar. O `.gitignore` não é uma ferramenta de anonimização nem apaga arquivos do histórico do Git.

Para executar os testes locais no Windows, sem consultar o portal:

```powershell
py -3.11 -m unittest test_tcp_monitor test_excel_open test_monitor_control test_integration_process
```

Os testes não substituem a validação dos botões e das macros no Excel nem a verificação de compatibilidade com o portal.

---

## Español

Automatización para consultar CE seleccionados en el portal TCP y seguir los resultados en Excel. El monitor puede realizar una consulta puntual o repetirlas a intervalos configurados.

### Componentes

- `tcp_monitor.py`: consulta, comparación de resultados y coordinación.
- `excel_open.py`: integración con Excel y protección frente a cambios concurrentes.
- `monitor_control.py`: señales de consulta inmediata y detención.
- `excel_buttons.bas` y `excel_workbook_events.txt`: código VBA para botones y panel.
- `modelo/tcp_monitoring_template.xlsx`: plantilla vacía; los rótulos y mensajes están en portugués.
- `test_*.py`: pruebas locales; `requirements.txt`: dependencias de Python.

### Requisitos

- Windows, Microsoft Excel de escritorio, Microsoft Edge, acceso al portal TCP y Python 3.11.
- Dependencias instaladas en Python 3.11:

```powershell
py -3.11 -m pip install -r requirements.txt
```

### Preparación

1. Copie `modelo/tcp_monitoring_template.xlsx` a la carpeta de los programas, conservando el nombre `tcp_monitoring_template.xlsx`. Use esa copia para sus datos.
2. En la hoja `Monitoramento`, introduzca cada CE como **texto de 15 dígitos** y seleccione `Sim` o `Não` en `Monitorar`.
3. En `Configurações`, elija un intervalo de 5, 15, 30, 45 o 60 minutos. Para el modo continuo, seleccione `Sim` en `Automação ativa`.

No introduzca datos reales en la plantilla versionada. `.gitignore` protege las hojas de trabajo, pero no impide una publicación accidental si se edita la plantilla versionada o se agregan archivos forzadamente.

### Consulta desde la terminal

En la carpeta de los programas:

```powershell
py -3.11 tcp_monitor.py
```

Para repetir consultas mientras la terminal permanezca abierta:

```powershell
py -3.11 tcp_monitor.py --loop
```

El modo continuo exige `Automação ativa` en `Sim`; use `Ctrl+C` para detenerlo desde la terminal. Use `--planilha` y `--dados` para indicar otra hoja u otro archivo de datos.

### Botones en Excel

El repositorio proporciona el **código de las macros**, no un archivo `.xlsm` listo. En la copia de trabajo, use **Guardar como** y seleccione un libro habilitado para macros, por ejemplo `tcp_monitoring.xlsm`, en la misma carpeta que los programas.

Confirme que las políticas de la organización permiten macros. Importe `excel_buttons.bas` desde el editor VBA (`Alt+F11`, `Archivo > Importar archivo`), cree dos botones en `Monitoramento` y asígneles `ConsultarAgora` e `IniciarMonitoramento`. Ejecute `PrepararPainel` una vez para crear el botón de parada y el panel; las celdas `E1:L4` deben estar vacías y no debe haber formas sobre las áreas de destino. Agregue los procedimientos de `excel_workbook_events.txt` a `ThisWorkbook`, conservando los eventos existentes, y guarde el libro.

- **Consultar agora:** ejecuta una consulta o solicita una inmediata al monitor activo. Si hay una consulta en curso, espera a que termine.
- **Iniciar monitoramento:** inicia el modo continuo; solo puede haber una ejecución a la vez.
- **Parar monitoramento:** después de confirmar, solicita cancelar la consulta actual; se descartan sus resultados.

El panel se actualiza mientras Excel está abierto. `Sem resposta` indica que el estado dejó de actualizarse; no confirma que el proceso funcione. Cambiar los archivos VBA externos no actualiza automáticamente el código importado.

### Funcionamiento y límites

- Solo se consultan los CE con `Sim`. Un CE sin selección queda pendiente sin impedir los demás.
- Con el libro abierto, guarde manualmente los resultados y la configuración que desee conservar. Si está cerrado, el monitor utiliza Excel para abrirlo y guardarlo temporalmente.
- El monitor puede continuar con Excel cerrado mientras Windows y el proceso sigan activos. No se inicia automáticamente tras apagar, reiniciar o cerrar la sesión.
- Los cambios concurrentes pueden impedir una actualización para proteger los datos. El JSON puede haberse actualizado y permitir la recuperación en el siguiente ciclo.
- `tcp_monitoring_data.json` guarda resultados para comparación y recuperación; `tcp_monitor.log` registra ejecuciones ocultas; `tcp_monitor_status.txt` alimenta el panel. Conserve esos archivos según sus necesidades.
- Cambios del portal, falta de acceso, bloqueos de Excel o políticas corporativas pueden impedir la ejecución. No eluda las políticas de seguridad; consulte con TI.

### Privacidad y pruebas

Los libros, JSON y registros pueden contener información comercial. Revise datos y metadatos antes de compartir. `.gitignore` no anonimiza ni elimina archivos del historial de Git.

Para ejecutar las pruebas locales en Windows sin consultar el portal:

```powershell
py -3.11 -m unittest test_tcp_monitor test_excel_open test_monitor_control test_integration_process
```

Las pruebas no sustituyen la validación de botones y macros en Excel ni la comprobación de compatibilidad con el portal.

---

## English

Automation for querying selected CEs on the TCP portal and tracking results in Excel. The monitor can perform a single query or repeat queries at configured intervals.

### Components

- `tcp_monitor.py`: querying, result comparison and coordination.
- `excel_open.py`: Excel integration and concurrent-change protection.
- `monitor_control.py`: immediate-query and stop signals.
- `excel_buttons.bas` and `excel_workbook_events.txt`: VBA code for buttons and the status panel.
- `modelo/tcp_monitoring_template.xlsx`: empty template; application labels and messages are in Portuguese.
- `test_*.py`: local tests; `requirements.txt`: Python dependencies.

### Requirements

- Windows, desktop Microsoft Excel, Microsoft Edge, access to the TCP portal and Python 3.11.
- Dependencies installed for Python 3.11:

```powershell
py -3.11 -m pip install -r requirements.txt
```

### Setup

1. Copy `modelo/tcp_monitoring_template.xlsx` into the program folder, keeping the name `tcp_monitoring_template.xlsx`. Use that copy for your data.
2. In the `Monitoramento` sheet, enter each CE as **15-digit text** and select `Sim` or `Não` in `Monitorar`.
3. In `Configurações`, select an interval of 5, 15, 30, 45 or 60 minutes. For continuous monitoring, set `Automação ativa` to `Sim`.

Do not enter real data in the versioned template. `.gitignore` protects working spreadsheets, but cannot prevent accidental publication if the tracked template is edited or files are force-added.

### Terminal use

In the program folder:

```powershell
py -3.11 tcp_monitor.py
```

To repeat queries while the terminal remains open:

```powershell
py -3.11 tcp_monitor.py --loop
```

Continuous mode requires `Automação ativa` to be `Sim`; use `Ctrl+C` to stop it from the terminal. Use `--planilha` and `--dados` to specify another workbook or data file.

### Excel buttons

The repository provides the **macro source code**, not a ready-made `.xlsm` workbook. In the working copy, use **Save As** and choose a macro-enabled workbook, for example `tcp_monitoring.xlsm`, in the same folder as the programs.

Confirm that your organization's policies permit macros. Import `excel_buttons.bas` in the VBA editor (`Alt+F11`, `File > Import File`), create two buttons on `Monitoramento` and assign `ConsultarAgora` and `IniciarMonitoramento`. Run `PrepararPainel` once to create the stop button and panel; cells `E1:L4` must be empty and no shapes may overlap the target areas. Add the procedures from `excel_workbook_events.txt` to `ThisWorkbook`, preserving existing events, and save the workbook.

- **Consultar agora:** performs a query or requests an immediate query from the active monitor. If a query is in progress, it waits until it finishes.
- **Iniciar monitoramento:** starts continuous mode; only one execution can run at a time.
- **Parar monitoramento:** after confirmation, requests cancellation of the current query; that cycle's results are discarded.

The panel updates while Excel is open. `Sem resposta` means the status stopped refreshing; it does not confirm that the process is running. Changing the external VBA files does not automatically update already imported code.

### Operation and limitations

- Only CEs marked `Sim` are queried. A CE without a selection remains pending without blocking the others.
- When the workbook is open, manually save results and settings you want to keep. When it is closed, the monitor uses Excel to open and save it temporarily.
- Monitoring can continue with Excel closed while Windows and the monitor process remain active. It does not start automatically after shutdown, restart or sign-out.
- Concurrent changes can prevent an update to protect data. The JSON may already have been updated and allow recovery on the next cycle.
- `tcp_monitoring_data.json` stores results for comparison and recovery; `tcp_monitor.log` records hidden executions; `tcp_monitor_status.txt` feeds the panel. Keep these files as needed.
- Portal changes, lack of access, Excel locks or corporate policies can prevent execution. Do not bypass security policies; consult IT.

### Privacy and tests

Workbooks, JSON and logs may contain business information. Review data and metadata before sharing. `.gitignore` does not anonymize data or remove files from Git history.

To run local tests on Windows without querying the portal:

```powershell
py -3.11 -m unittest test_tcp_monitor test_excel_open test_monitor_control test_integration_process
```

These tests do not replace validation of buttons and macros in Excel or portal compatibility checks.
