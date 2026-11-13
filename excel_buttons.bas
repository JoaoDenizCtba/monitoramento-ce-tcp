Attribute VB_Name = "MonitoramentoTCP"
Option Explicit

Private ProximaAtualizacao As Date

Public Sub ConsultarAgora()
    IniciarMonitor ""
End Sub

Public Sub IniciarMonitoramento()
    IniciarMonitor "--loop"
End Sub

Public Sub PararMonitoramento()
    If MsgBox("Interromper o monitoramento e a consulta em andamento?", _
              vbYesNo + vbQuestion, "Monitoramento TCP") = vbYes Then
        IniciarMonitor "--parar"
    End If
End Sub

Private Sub IniciarMonitor(ByVal acao As String)
    Dim pasta As String
    Dim script As String
    Dim comando As String

    On Error GoTo Falha
    pasta = ThisWorkbook.Path
    If Len(pasta) = 0 Or InStr(1, pasta, "://", vbTextCompare) > 0 Then
        MsgBox "Salve a planilha em uma pasta local junto com o script antes de iniciar.", vbExclamation
        Exit Sub
    End If
    script = pasta & Application.PathSeparator & "tcp_monitor.py"
    If Len(Dir$(script)) = 0 Then
        MsgBox "O arquivo tcp_monitor.py não foi encontrado na pasta da planilha.", vbExclamation
        Exit Sub
    End If
    comando = "py.exe -3.11 """ & script & """ --planilha """ & ThisWorkbook.FullName & """ --oculto " & acao
    CreateObject("WScript.Shell").Run comando, 0, False
    Exit Sub

Falha:
    MsgBox "Não foi possível iniciar a consulta: " & Err.Description, vbExclamation
End Sub

Public Sub PrepararPainel()
    Dim aba As Worksheet
    Dim forma As Shape
    Dim botaoCriado As Shape
    Dim statusCriado As Shape
    Dim areaParar As Range
    Dim areaStatus As Range
    Dim detalhe As String

    On Error GoTo Falha
    Set aba = ThisWorkbook.Worksheets("Monitoramento")
    Set areaParar = aba.Range("E2:F3")
    Set areaStatus = aba.Range("H1:L4")
    If Application.WorksheetFunction.CountA(aba.Range("E1:L4")) > 0 Then
        Err.Raise vbObjectError + 1, , "As células E1:L4 precisam estar vazias para criar o painel."
    End If
    For Each forma In aba.Shapes
        If forma.Name = "PararMonitorTCP" Or forma.Name = "StatusMonitorTCP" Then
            Err.Raise vbObjectError + 2, , "O painel já existe; nenhuma forma foi criada."
        End If
        If Sobrepoe(forma, areaParar) Or Sobrepoe(forma, areaStatus) Then
            Err.Raise vbObjectError + 3, , "Há uma forma na área do novo painel; nenhuma forma foi criada."
        End If
    Next forma

    Set botaoCriado = aba.Shapes.AddShape(msoShapeRoundedRectangle, areaParar.Left, areaParar.Top, _
                                          areaParar.Width, areaParar.Height)
    botaoCriado.Name = "PararMonitorTCP"
    botaoCriado.TextFrame.Characters.Text = "Parar monitoramento"
    botaoCriado.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PararMonitoramento"
    Set statusCriado = aba.Shapes.AddShape(msoShapeRoundedRectangle, areaStatus.Left, areaStatus.Top, _
                                           areaStatus.Width, areaStatus.Height)
    statusCriado.Name = "StatusMonitorTCP"
    statusCriado.TextFrame.Characters.Text = "Status: aguardando atualização"
    IniciarAtualizacaoStatus
    Exit Sub

Falha:
    detalhe = Err.Description
    If Not statusCriado Is Nothing Then statusCriado.Delete
    If Not botaoCriado Is Nothing Then botaoCriado.Delete
    MsgBox "Não foi possível preparar o painel: " & detalhe, vbExclamation
End Sub

Private Function Sobrepoe(ByVal forma As Shape, ByVal area As Range) As Boolean
    Sobrepoe = forma.Left < area.Left + area.Width And forma.Left + forma.Width > area.Left _
        And forma.Top < area.Top + area.Height And forma.Top + forma.Height > area.Top
End Function

Public Sub IniciarAtualizacaoStatus()
    CancelarAtualizacaoStatus
    AtualizarStatus
End Sub

Public Sub CancelarAtualizacaoStatus()
    If ProximaAtualizacao > Now Then
        On Error GoTo Falha
        Application.OnTime ProximaAtualizacao, NomeMacroStatus(), , False
    End If
    ProximaAtualizacao = 0
    Exit Sub
Falha:
    MsgBox "Não foi possível cancelar a atualização do status: " & Err.Description, vbExclamation
End Sub

Private Function NomeMacroStatus() As String
    NomeMacroStatus = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AtualizarStatus"
End Function

Public Sub AtualizarStatus()
    Dim aba As Worksheet
    Dim forma As Shape
    Dim arquivo As String
    Dim fluxo As Object
    Dim caminho As String
    Dim situacao As String
    Dim ultima As String
    Dim proxima As String
    Dim texto As String
    Dim estavaSalva As Boolean

    ProximaAtualizacao = 0
    On Error GoTo Falha
    Set aba = ThisWorkbook.Worksheets("Monitoramento")
    Set forma = aba.Shapes("StatusMonitorTCP")
    arquivo = ThisWorkbook.Path & Application.PathSeparator & "tcp_monitor_status.txt"
    texto = "Monitoramento: Parado" & vbLf & "Último ciclo: -" & vbLf & "Próximo ciclo: -"
    If Len(Dir$(arquivo)) > 0 Then
        Set fluxo = CreateObject("Scripting.FileSystemObject").OpenTextFile(arquivo, 1, False, -1)
        caminho = fluxo.ReadLine
        situacao = fluxo.ReadLine
        ultima = fluxo.ReadLine
        proxima = fluxo.ReadLine
        fluxo.Close
        If StrComp(caminho, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            If situacao <> "Parado" And situacao <> "Erro" And DateDiff("s", FileDateTime(arquivo), Now) > 30 Then
                situacao = "Sem resposta"
                proxima = "-"
            End If
            texto = "Monitoramento: " & situacao & vbLf & "Último ciclo: " & ultima & vbLf & "Próximo ciclo: " & proxima
        End If
    End If
    If forma.TextFrame.Characters.Text <> texto Then
        estavaSalva = ThisWorkbook.Saved
        forma.TextFrame.Characters.Text = texto
        If estavaSalva Then ThisWorkbook.Saved = True
    End If
    ProximaAtualizacao = Now + TimeSerial(0, 0, 10)
    Application.OnTime ProximaAtualizacao, NomeMacroStatus()
    Exit Sub
Falha:
    If Not fluxo Is Nothing Then fluxo.Close
    MsgBox "Não foi possível atualizar o status: " & Err.Description, vbExclamation
End Sub
