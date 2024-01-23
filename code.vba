Option Explicit
Public SapGuiAuto   As Variant
Public SAPApp       As Variant
Public SAPCon       As Variant
Public session      As Variant
Public Connection   As Variant
Public WScript      As Variant
Public ArquivoComando       As Workbook
Public ArquivoKanban        As Workbook
Public PlanilhaComando      As Worksheet
Public Kanban               As Worksheet
Public PlanilhaExportada    As Worksheet
Public PlanilhaDatas        As Worksheet
Public ListaPEP             As Worksheet
Public SystemStatus         As String
Public StatusCor            As String
Public DiasMontagemA        As Integer
Public DiasMontagemB        As Integer
Public DiasMontagemC        As Integer

Public CNCfeito As Boolean
Public CortesFeito As Boolean
Sub DeclararVariaveis()

    Set ArquivoComando = ActiveWorkbook
    Set ArquivoKanban = Workbooks.Open(ThisWorkbook.Path & "\Kanban.XLSM")
    Set PlanilhaComando = ArquivoComando.Worksheets("Comando")
    
    Set ListaPEP = ArquivoComando.Worksheets("Lista PEP")
    Set Kanban = ArquivoKanban.Worksheets("Tampa Usinada")
    
    SystemStatus = "CONF"

    DiasMontagemA = -5
    DiasMontagemB = -7
    DiasMontagemC = -8
End Sub

Sub ConectarSAP()
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    
    If Not IsObject(Application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject Application, "on"
    End If
End Sub

Option Explicit
Function ContarLinhas(ws_temp As Worksheet)
    ContarLinhas = ws_temp.UsedRange.Rows.Count
End Function
Function ContarColunas(ws_temp As Worksheet)
    ContarColunas = ws_temp.UsedRange.Columns.Count
End Function
Sub Estilizar(ws As Worksheet)
    ws.Range("A1:V2").Interior.ColorIndex = 48
End Sub
Sub AjustarColunas(ws As Worksheet)
    ws.Cells.Select
    ws.Cells.EntireColumn.AutoFit
    ws.Selection.HorizontalAlignment = xlCenter
End Sub
Sub ClearFormat()
    Set ArquivoPrincipal = Workbooks("Analise Diversos Montagem.XLSM")
    Set PlanilhaPrincipal = ArquivoPrincipal.Worksheets("Comando")
    PlanilhaPrincipal.Range("A3:N1000").ClearFormats
    PlanilhaPrincipal.Range("A3:N1000").Value = ""
End Sub
Sub Main()
    Call DeclararVariaveis
    Call ConectarSAP
    Call FomatarPlanilhas
    
    Call LimparLista
    Call ImportarDadosKanban
    Call EntrarTelaMontagemSAP
    Call ImportarMontagem
    Call ExaminarConsulta_InserirNaLista
    
    Call EntrarCOOIS
    Call ColorirPlanilha
    
    Call LimparCalendario
    Call FomatarPlanilhas
    Call ImportarListaPEP
End Sub

Sub EntrarTelaMontagemSAP()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCN43N"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/MONT_REQ"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").SetFocus
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 9
End Sub

Sub LimparLista()
    ListaPEP.Cells.ClearContents
    ListaPEP.Columns("A").ClearFormats
End Sub

Sub LimparCalendario()
    PlanilhaComando.Range("A3:N1000").ClearContents
    PlanilhaComando.Range("A3:N1000").ClearFormats
End Sub

Sub ImportarDadosKanban()
    ArquivoKanban.RefreshAll
    QuantidadePEPs = ContarLinhas(Kanban)
    
    For Row = 1 To QuantidadePEPs
        DataFimBase = Kanban.Cells(Row, 11).Value
        DataFimReal = Kanban.Cells(Row, 12).Value
        
        If DataFimReal = "" Then
            NumeroPEP = Kanban.Cells(Row, 9).Value
            InserirPepNaLista NumeroPEP, DataFimBase
        End If
    Next Row
    
    ArquivoKanban.Close SaveChanges:=False
End Sub

Sub InserirPepNaLista(NumeroPEP, DataFimBase)
    LinhaAtual = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    LinhaAtualDeBaixo = LinhaAtual + 1
    
    ListaPEP.Cells(LinhaAtualDeBaixo, 1) = NumeroPEP
    ListaPEP.Cells(LinhaAtualDeBaixo, 2) = DataFimBase
End Sub

Sub ImportarMontagem()
    LinhaAtual = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    TotalPEPs = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    For Row = 2 To TotalPEPs
        NumeroPEP = ListaPEP.Cells(Row, 1).Value
        
        InserirPEPConsulta NumeroPEP
    Next Row
    
    EnviarConsulta
End Sub

Sub RemoverBacklog()
    TotalLinhas = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To TotalLinhas
        DataFimReal = ListaPEP.Cells(Row, 4).Value
        DataHoje = Date
        
        If DataFimReal < DataHoje Then
            With ListaPEP
                .Cells(Row, 1).Value = ""
                .Cells(Row, 2).Value = ""
                .Cells(Row, 3).Value = ""
                .Cells(Row, 4).Value = ""
            End With
        End If
    Next Row
End Sub

Sub InserirPEPConsulta(NumeroPEP)
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 1
    session.findById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").press
    
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = NumeroPEP
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
    
    PosicaoScroll = session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = PosicaoScroll + 3
End Sub

Sub EnviarConsulta()
    session.findById("wnd[1]").Close
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Sub ExaminarConsulta_InserirNaLista()
    TotalRows = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 0 To TotalRows
        Montagem = session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").GetCellValue(Row, "VERNA")
        PEP = session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").GetCellValue(Row, "POSKI")
    
        InserirMontagemDataNaLista PEP, Montagem
    Next Row
End Sub

Sub InserirMontagemDataNaLista(PEPConsulta, Montagem)
    TotalPEPs = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    For Row = 2 To TotalPEPs
    
        PEPLista = ListaPEP.Cells(Row, 1).Value
        DataFimBase = ListaPEP.Cells(Row, 2).Value
        
        If PEPLista = PEPConsulta Then
            Montagem = FormatarLetraMontagem(Montagem)
            DataFimReal = SubtrairDiasUteis(Montagem, DataFimBase)
            
            ListaPEP.Cells(Row, 3).Value = Montagem
            ListaPEP.Cells(Row, 4).Value = DataFimReal
            
            Exit For
        End If
    Next Row
End Sub

Function FormatarLetraMontagem(TextoMontagem)
    TextoMontagem = Replace(TextoMontagem, "Montagem Final ", "")
    TextoMontagem = Replace(TextoMontagem, Chr(34), "")
    FormatarLetraMontagem = TextoMontagem
End Function

Function SubtrairDiasUteis(Montagem, DataFimBase)
    Select Case Montagem
        Case Is = "A"
            DiasAnteriores = -12
        Case Is = "B"
            DiasAnteriores = -14
        Case Is = "C"
            DiasAnteriores = -15
    End Select
    
    SubtrairDiasUteis = WorksheetFunction.WorkDay(DataFimBase, DiasAnteriores)
End Function

Sub EntrarCOOIS()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").Text = ""
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").SetFocus
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOO000"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-LOW").Text = "108"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_DISPO-HIGH").Text = "109"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").SetFocus
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 6
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/OPCALDPREP"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").SetFocus
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 11
End Sub

Sub ColorirPlanilha()
    TotalPEPs = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    For Row = 2 To TotalPEPs
        PEP = ListaPEP.Cells(Row, 1).Value
        
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").Text = PEP
        session.findById("wnd[0]/tbar[1]/btn[8]").press
            
        If ExisteConsulta() Then
            ListaPEP.Cells(Row, 1).Interior.ThemeColor = StatusComponentes()
            
            session.findById("wnd[0]/tbar[0]/btn[15]").press
        End If
    Next Row
End Sub

Function ExisteConsulta()
    If Not session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell", False) Is Nothing Then
        ExisteConsulta = True
        Exit Function
    End If
    
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    ExisteConsulta = False
End Function

Function StatusComponentes()
    TotalLinhas = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").VisibleRowcount
    StatusConferido = "CONF"
    
    StatusComponentes = xlThemeColorAccent6
    
    CortesFeito = True
    CNCfeito = True
    
    For Linha = 0 To TotalLinhas - 1
        ValorStatus = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(Linha, "VSTTXT") '
        ValorTexto = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(Linha, "LTXA1")
        
        If ValorTexto = "FAZER PROGRAMA CNC" Then
            If InStr(ValorStatus, StatusConferido) = 0 Then
                CNCfeito = False
                StatusComponentes = xlThemeColorAccent5 'azul
            End If
        End If
        
        If CNCfeito = True Then
            If InStr(ValorTexto, "CORTAR") > 0 Then
                If InStr(ValorStatus, StatusConferido) = 0 Then
                    StatusComponentes = xlThemeColorAccent4 'amarelo
                End If
            End If
        End If
        
    Next Linha
End Function

Sub FomatarPlanilhas()
    Call SetarDatasCabeçalho(PlanilhaComando)
    Call FormatarCabeçalho(PlanilhaComando)
End Sub

Sub SetarDatasCabeçalho(Planilha As Worksheet)
    For Coluna = 2 To 14
        Planilha.Cells(2, Coluna) = WorksheetFunction.WorkDay(Date, Coluna - 2)
        Planilha.Range("A3:N1000") = ""
        Planilha.Range("A3:N1000").ClearFormats
    Next Coluna
End Sub

Sub FormatarCabeçalho(Planilha)
    With Planilha
        .Cells.EntireColumn.AutoFit
        .Range("A1:N1").ColumnWidth = 15
        .Range("A1:N999").HorizontalAlignment = xlCenter
    End With
End Sub

Sub ImportarListaPEP()
    TotalPEPs = ListaPEP.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To TotalPEPs
        PEP = ListaPEP.Cells(Row, 1).Value
        DataFimReal = ListaPEP.Cells(Row, 4).Value
        
        If DataEstaNoBacklog(DataFimReal) Then
            ListaPEP.Cells(Row, 1).Copy
            PreencherPEPBacklog
        Else
            ListaPEP.Cells(Row, 1).Copy
            CompararDatasCabeçalho DataFimReal
        End If
        
    Next Row
End Sub

Sub CompararDatasCabeçalho(DataFimReal)
    For ColunaData = 2 To 14
            
            DataCabeçalho = PlanilhaComando.Cells(2, ColunaData)
                
            If DataFimReal = DataCabeçalho Then
                PreencherPEP ColunaData
            End If
            
        Next ColunaData
End Sub

Sub PreencherPEP(ColunaData)
    LinhaAtualPEP = PlanilhaComando.Cells(Rows.Count, ColunaData).End(xlUp).Row
    ProximaLinhaDeBaixo = LinhaAtualPEP + 1
    
    PlanilhaComando.Cells(ProximaLinhaDeBaixo, ColunaData).PasteSpecial Paste:=xlPasteFormats
    PlanilhaComando.Cells(ProximaLinhaDeBaixo, ColunaData).PasteSpecial Paste:=xlPasteValues
End Sub

Function DataEstaNoBacklog(DataFinal)
    DataEstaNoBacklog = False
    
    DataHoje = Date
    If DataFinal < DataHoje Then
        DataEstaNoBacklog = True
    End If
End Function

Sub PreencherPEPBacklog()
    ColunaBackLog = 1
    
    LinhaAtualPreenchida = PlanilhaComando.Cells(Rows.Count, ColunaBackLog).End(xlUp).Row
    ProximaLinhaDeBaixo = LinhaAtualPreenchida + 1
    
    PlanilhaComando.Cells(ProximaLinhaDeBaixo, ColunaBackLog).PasteSpecial Paste:=xlPasteFormats
    PlanilhaComando.Cells(ProximaLinhaDeBaixo, ColunaBackLog).PasteSpecial Paste:=xlPasteValues
End Sub

Sub AtualizarCalendario()
    Call DeclararVariaveis
    Call ConectarSAP
    For ColunaData = 2 To 14
        TotalLinhasPEP = PlanilhaComando.Cells(Rows.Count, ColunaData).End(xlUp).Row
        
        For LinhaPEP = 3 To TotalLinhasPEP
            CelulaPEP = PlanilhaComando.Cells(LinhaPEP, ColunaData)
            PEP = PlanilhaComando.Cells(LinhaPEP, ColunaData).Value
            
            ' TODO: SE FOR COR BRANCA DA ERRO
            If PlanilhaComando.Cells(LinhaPEP, ColunaData).Interior.ThemeColor <> xlThemeColorAccent6 Then
               
                EntrarCOOIS
                AcessarPEP PEP
                
                 If ExisteConsulta() Then
                    PlanilhaComando.Cells(LinhaPEP, ColunaData).Interior.ThemeColor = StatusComponentes()
                    
                    session.findById("wnd[0]/tbar[0]/btn[15]").press
                End If
                
            End If
            
        Next LinhaPEP
    Next ColunaData
End Sub

Function AcessarPEP(PEP)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PROJN-LOW").Text = PEP
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
End Function


