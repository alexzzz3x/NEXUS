On Error Resume Next
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

' Pressiona btn[12] para retornar/ativar a barra de comando e digita /nIQ09
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIQ09"
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Then
   Err.Clear
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00027"
End If
On Error GoTo 0

session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press

WScript.Sleep 800

' Trata eventuais popups informativos ou avisos
If session.Children.Count > 1 Then
    On Error Resume Next
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 300
    On Error GoTo 0
End If

Dim eq, desc, st, mat, ser, centro, dep, tipoEstq, stUser
eq = "" : desc = "" : st = "" : mat = "" : ser = "" : centro = "" : dep = "" : tipoEstq = "" : stUser = ""

' 1. Verifica se retornou grade ALV
Dim grid
Set grid = Nothing
On Error Resume Next
Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
If Not grid Is Nothing Then
    ' Função para coletar todos os status da linha (sistema e usuário)
    Function ColetarStatusLinhaALV(g, rowIdx)
        Dim stColDict, cI, colN, valPiece, words, w, resSt
        Set stColDict = CreateObject("Scripting.Dictionary")
        
        ' 1. Lista de campos padrão de status do SAP ALV
        Dim knownStatusCols, kCol
        knownStatusCols = Array("STTXT", "ASTTX", "STEXT", "TXT04", "TXT30", "ANWS", "STTXU", "SYSST", "USRST", "STSYN", "STUSN", "STAT", "STATUS", "ESTAT", "STBLA", "STTXT_EXT", "STTXT_INT", "R_STATUS")
        For Each kCol In knownStatusCols
            valPiece = ""
            On Error Resume Next
            valPiece = Trim(g.GetCellValue(rowIdx, kCol))
            On Error GoTo 0
            If valPiece <> "" And UCase(valPiece) <> "STATUS" And UCase(valPiece) <> "STAT" And Len(valPiece) <= 40 Then
                words = Split(valPiece, " ")
                For Each w In words
                    w = Trim(w)
                    If w <> "" And Len(w) <= 15 And Not stColDict.Exists(UCase(w)) Then
                        stColDict.Add UCase(w), w
                    End If
                Next
            End If
        Next
        
        ' 2. Varre todas as colunas visíveis da grade buscando colunas de status ou códigos conhecidos
        On Error Resume Next
        Dim allCols
        allCols = g.ColumnOrder
        If IsArray(allCols) Then
            For cI = 0 To UBound(allCols)
                colN = Trim(allCols(cI))
                If colN <> "SERNR" And colN <> "MATNR" And colN <> "EQUNR" And colN <> "MAKTX" And colN <> "SHTXT" And colN <> "EQKTX" And colN <> "WERKS" And colN <> "SWERK" And colN <> "LGORT" And colN <> "LAGER" And colN <> "LBBSA" Then
                    valPiece = Trim(g.GetCellValue(rowIdx, colN))
                    If valPiece <> "" Then
                        Dim uVal
                        uVal = UCase(valPiece)
                        If InStr(UCase(colN), "ST") > 0 Or InStr(UCase(colN), "STAT") > 0 Or InStr(UCase(colN), "TXT") > 0 Or InStr(UCase(colN), "ANW") > 0 Or InStr(UCase(colN), "USR") > 0 Or InStr(UCase(colN), "SYS") > 0 Or uVal = "LIDI" Or uVal = "DISP" Or uVal = "ELAB" Or uVal = "MONT" Or uVal = "ESTO" Or uVal = "BLOQ" Or uVal = "SUCA" Or uVal = "CONF" Or uVal = "INST" Or uVal = "AVLB" Or uVal = "DESM" Or uVal = "MOVEL" Then
                            If uVal <> "STATUS" And uVal <> "STAT" And uVal <> "EQUIPAMENTO" And uVal <> "MATERIAL" Then
                                words = Split(valPiece, " ")
                                For Each w In words
                                    w = Trim(w)
                                    If w <> "" And Len(w) <= 15 And Not stColDict.Exists(UCase(w)) Then
                                        stColDict.Add UCase(w), w
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            Next
        End If
        On Error GoTo 0
        
        resSt = ""
        For Each w In stColDict.Keys
            If resSt = "" Then
                resSt = stColDict(w)
            Else
                resSt = resSt & " " & stColDict(w)
            End If
        Next
        ColetarStatusLinhaALV = Trim(resSt)
    End Function

    If grid.RowCount > 1 Then
        Dim r, rEq, rDesc, rSt, rMat, rSer, rCentro, rDep, rTipo
        Dim reCleanRow
        Set reCleanRow = CreateObject("VBScript.RegExp")
        reCleanRow.Pattern = "\s+"
        reCleanRow.Global = True

        For r = 0 To grid.RowCount - 1
            rEq = "" : rDesc = "" : rSt = "" : rMat = "" : rSer = "" : rCentro = "" : rDep = "" : rTipo = ""
            On Error Resume Next
            rEq = Trim(grid.GetCellValue(r, "EQUNR"))
            rDesc = Trim(grid.GetCellValue(r, "MAKTX"))
            If rDesc = "" Then rDesc = Trim(grid.GetCellValue(r, "SHTXT"))
            If rDesc = "" Then rDesc = Trim(grid.GetCellValue(r, "EQKTX"))

            rSt = ColetarStatusLinhaALV(grid, r)
            rSt = reCleanRow.Replace(rSt, " ")

            rMat = Trim(grid.GetCellValue(r, "MATNR"))
            rSer = Trim(grid.GetCellValue(r, "SERNR"))
            rCentro = Trim(grid.GetCellValue(r, "WERKS"))
            If rCentro = "" Then rCentro = Trim(grid.GetCellValue(r, "SWERK"))
            If rCentro = "" Then rCentro = "Sem informacao"
            rDep = Trim(grid.GetCellValue(r, "LGORT"))
            If rDep = "" Then rDep = Trim(grid.GetCellValue(r, "LAGER"))
            If rDep = "" Then rDep = "Sem informacao"
            rTipo = Trim(grid.GetCellValue(r, "LBBSA"))
            On Error GoTo 0
            
            WScript.Echo "MEDIDOR_INFO:EQ=" & rEq & "|DESC=" & rDesc & "|STATUS=" & rSt & "|MAT=" & rMat & "|SERIAL=" & rSer & "|CENTRO=" & rCentro & "|DEP=" & rDep & "|TIPO=" & rTipo
        Next
        WScript.Quit 0
    ElseIf grid.RowCount = 1 Then
        eq = Trim(grid.GetCellValue(0, "EQUNR"))
        desc = Trim(grid.GetCellValue(0, "MAKTX"))
        If desc = "" Then desc = Trim(grid.GetCellValue(0, "SHTXT"))
        st = ColetarStatusLinhaALV(grid, 0)
        mat = Trim(grid.GetCellValue(0, "MATNR"))
        ser = Trim(grid.GetCellValue(0, "SERNR"))
        centro = Trim(grid.GetCellValue(0, "WERKS"))
        If centro = "" Then centro = Trim(grid.GetCellValue(0, "SWERK"))
        dep = Trim(grid.GetCellValue(0, "LGORT"))
        If dep = "" Then dep = Trim(grid.GetCellValue(0, "LAGER"))

        ' Se houver 1 resultado, dá duplo clique para abrir a tela de Detalhe nº série
        grid.doubleClickCurrentCell
        WScript.Sleep 700
    End If
End If
On Error GoTo 0

' 2. Seleciona a aba Dds.série (tabpT\06) para carregar os campos de Material, Serial, Centro e Depósito
Sub SelecionarAbaSerie(obj)
    On Error Resume Next
    Dim i, child, cId, cType, cText
    For i = 0 To obj.Children.Count - 1
        Set child = obj.Children(CInt(i))
        If Not child Is Nothing Then
            cId = UCase(child.Id)
            cType = child.Type
            cText = UCase(child.Text)
            If cType = "GuiTab" And (InStr(cId, "TABPT\06") > 0 Or InStr(cId, "TABPT/06") > 0 Or InStr(cText, "SÉRIE") > 0 Or InStr(cText, "SERIE") > 0 Or InStr(cText, "DDS") > 0) Then
                child.select
                Exit Sub
            End If
            If child.Children.Count > 0 Then
                SelecionarAbaSerie child
            End If
        End If
    Next
End Sub

On Error Resume Next
session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06").select
WScript.Sleep 600
If Err.Number <> 0 Then
    Err.Clear
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/tabsTABSTRIP/tabpT\06").select
    WScript.Sleep 600
End If
If Err.Number <> 0 Then
    Err.Clear
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLIEE0:0100/tabsTABSTRIP/tabpT\06").select
    WScript.Sleep 600
End If
If Err.Number <> 0 Then
    Err.Clear
    SelecionarAbaSerie session.findById("wnd[0]/usr")
    WScript.Sleep 600
End If
On Error GoTo 0

' 3. Leitura direta dos campos padrão do SAP GUI
On Error Resume Next
If eq = "" Then eq = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/txtITOB-EQUNR").text)
If eq = "" Then eq = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0152/txtRM63E-EQUNR").text)
If eq = "" Then eq = Trim(session.findById("wnd[0]/usr/txtRM63E-EQUNR").text)
If eq = "" Then eq = Trim(session.findById("wnd[0]/usr/txtITOB-EQUNR").text)

If desc = "" Then desc = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/txtITOB-SHTXT").text)
If desc = "" Then desc = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0152/txtRM63E-SHTXT").text)
If desc = "" Then desc = Trim(session.findById("wnd[0]/usr/txtRM63E-SHTXT").text)
If desc = "" Then desc = Trim(session.findById("wnd[0]/usr/txtITOB-SHTXT").text)

If st = "" Then st = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/txtRM63E-STTXT").text)
If st = "" Then st = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0152/txtRM63E-STTXT").text)
If st = "" Then st = Trim(session.findById("wnd[0]/usr/txtRM63E-STTXT").text)
If st = "" Then st = Trim(session.findById("wnd[0]/usr/txtITOB-STTXT").text)

' Le status de usuario se ainda nao preenchido
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/txtRM63E-ASTTX").text)
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0152/txtRM63E-ASTTX").text)
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/txtRM63E-ASTTX").text)
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/txtRM63E-STEXT").text)
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0152/txtRM63E-STEXT").text)
If stUser = "" Then stUser = Trim(session.findById("wnd[0]/usr/txtRM63E-STEXT").text)
If stUser <> "" And InStr(UCase(stUser), "STATUS") = 0 Then
    If st <> "" Then
        If InStr(st, stUser) = 0 Then st = st & " " & stUser
    Else
        st = stUser
    End If
End If

If mat = "" Then mat = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-MATNR").text)
If mat = "" Then mat = Trim(session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-MATNR").text)

If ser = "" Then ser = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/txtITOB-SERNR").text)
If ser = "" Then ser = Trim(session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/txtITOB-SERNR").text)

If centro = "" Then centro = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-SWERK").text)
If centro = "" Then centro = Trim(session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-SWERK").text)

If dep = "" Then dep = Trim(session.findById("wnd[0]/usr/subSUB_ALL:SAPLITO0:0100/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-LAGER").text)
If dep = "" Then dep = Trim(session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_BLOC:SAPLITO0:0120/ctxtITOB-LAGER").text)
On Error GoTo 0

' 4. Varredura recursiva de suporte dos controles da tela de Detalhe de Equipamento
Sub VarreTelaEquipamento(obj)
    On Error Resume Next
    Dim i, child, cId, cType, t, uT
    For i = 0 To obj.Children.Count - 1
        Set child = obj.Children(CInt(i))
        If Not child Is Nothing Then
            cId = UCase(child.Id)
            cType = child.Type
            t = Trim(child.Text)
            uT = UCase(t)

            ' CRÍTICO: Lê APENAS campos de texto e caixas de dados, NUNCA rótulos (GuiLabel), botões ou abas!
            If (cType = "GuiTextField" Or cType = "GuiCTextField" Or cType = "GuiComboBox") And t <> "" Then
                If InStr(cId, "EQUNR") > 0 And eq = "" And uT <> "EQUIPAMENTO" Then
                    eq = t
                End If

                If (InStr(cId, "SHTXT") > 0 Or InStr(cId, "EQKTX") > 0 Or InStr(cId, "MAKTX") > 0) And desc = "" And InStr(uT, "DENOMINA") = 0 Then
                    desc = t
                End If

                ' Detecção de Status: ID contendo status OU valor correspondente a status do SAP (DISP, LIDI, etc.)
                Dim isStatusControl
                isStatusControl = False
                If InStr(cId, "STTXT") > 0 Or InStr(cId, "ASTTX") > 0 Or InStr(cId, "STEXT") > 0 Or InStr(cId, "STATUS") > 0 Or InStr(cId, "STAT") > 0 Or InStr(cId, "ANWS") > 0 Or InStr(cId, "TXT04") > 0 Or InStr(cId, "TXT30") > 0 Or InStr(cId, "SYSST") > 0 Or InStr(cId, "USRST") > 0 Then
                    isStatusControl = True
                End If

                If uT = "DISP" Or uT = "LIDI" Or uT = "INST" Or uT = "MONT" Or uT = "MOVEL" Or uT = "MÓVEL" Or uT = "AVLB" Or uT = "DESM" Or uT = "BLOQ" Or uT = "SUCA" Or uT = "CONF" Or uT = "ELAB" Or uT = "ESTO" Then
                    isStatusControl = True
                End If

                If isStatusControl And uT <> "STATUS" And uT <> "STAT" And uT <> "EQUIPAMENTO" And uT <> "MATERIAL" And Len(t) <= 40 Then
                    If st = "" Then
                        st = t
                    ElseIf InStr(st, t) = 0 Then
                        st = st & " " & t
                    End If
                End If

                If InStr(cId, "MATNR") > 0 And mat = "" And uT <> "MATERIAL" Then
                    mat = t
                End If

                If InStr(cId, "SERNR") > 0 And ser = "" And InStr(uT, "SERIE") = 0 And InStr(uT, "SÉRIE") = 0 Then
                    ser = t
                End If

                If (InStr(cId, "SWERK") > 0 Or InStr(cId, "WERKS") > 0) And centro = "" And uT <> "CENTRO" Then
                    centro = t
                End If

                If (InStr(cId, "LAGER") > 0 Or InStr(cId, "LGORT") > 0) And dep = "" And InStr(uT, "DEPOSITO") = 0 And InStr(uT, "DEPÓSITO") = 0 Then
                    dep = t
                End If

                If InStr(cId, "LBBSA") > 0 And tipoEstq = "" And InStr(uT, "TIPO") = 0 Then
                    tipoEstq = t
                End If
            End If

            If child.Children.Count > 0 Then
                VarreTelaEquipamento child
            End If
        End If
    Next
End Sub

On Error Resume Next
VarreTelaEquipamento session.findById("wnd[0]")
On Error GoTo 0

' Junta status do usuário se houver e não estiver incluso
If stUser <> "" Then
    If st <> "" Then
        If InStr(st, stUser) = 0 Then st = st & " " & stUser
    Else
        st = stUser
    End If
End If

' Normaliza espaços duplicados no status (ex: "LIDI               DISP" -> "LIDI DISP")
Dim reClean
Set reClean = CreateObject("VBScript.RegExp")
reClean.Pattern = "\s+"
reClean.Global = True
st = reClean.Replace(st, " ")

If centro = "" Then centro = "Sem informacao"
If dep = "" Then dep = "Sem informacao"

If ser <> "" Or eq <> "" Or st <> "" Or mat <> "" Then
    WScript.Echo "MEDIDOR_INFO:EQ=" & eq & "|DESC=" & desc & "|STATUS=" & st & "|MAT=" & mat & "|SERIAL=" & ser & "|CENTRO=" & centro & "|DEP=" & dep & "|TIPO=" & tipoEstq
End If


