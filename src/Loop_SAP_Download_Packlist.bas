Option Explicit

' ============================================================
' SAP PACKING LIST DOWNLOADER
' ============================================================
'
' Purpose:
'   Automatically download Packing List attachments from SAP
'   (transaction VT03N) based on a list of Transport Numbers
'   in Excel.
'
' Technologies:
'   VBA
'   SAP GUI Scripting
'   Microsoft Excel
'
' ============================================================

Public Sub Loop_SAP_Download_Packlist()

    On Error GoTo ErroHandler

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    
    Dim ws As Worksheet
    Dim tknum As String
    Dim tknumOriginal As String
    Dim Caminho As String
    
    Dim tabela As Object
    Dim i As Long
    Dim titulo As String
    Dim Encontrou As Boolean
    
    Dim Linha As Long
    Dim ultimaLinha As Long

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Set ws = ThisWorkbook.Sheets("PACKINGLIST")     ' Corrected sheet name
    
    Caminho = Environ("USERPROFILE") & "\Downloads\Intercoout\"
    
    ' Create folder if it does not exist
    If Dir(Caminho, vbDirectory) = "" Then
        MkDir Caminho
    End If

    ' --------------------------------------------------------
    ' Connect to SAP GUI
    ' --------------------------------------------------------
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    
    If session Is Nothing Then
        MsgBox "Sessão SAP não encontrada.", vbCritical
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' Find last filled row
    ' --------------------------------------------------------
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Linha = 2

    ' --------------------------------------------------------
    ' Main Loop
    ' --------------------------------------------------------
    
    Do While Linha <= ultimaLinha And Trim(ws.Cells(Linha, 1).Value) <> ""
        
        Encontrou = False
        tknumOriginal = Trim(CStr(ws.Cells(Linha, 1).Value))
        tknum = tknumOriginal
        
        ' Clean invalid characters for filename
        tknum = Replace(tknum, "/", "-")
        tknum = Replace(tknum, "\", "-")
        
        ' ====================================================
        ' SAP Process
        ' ====================================================
        
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVT03N"
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").Text = tknumOriginal
        session.findById("wnd[0]").sendVKey 0
        
        ' Open GOS Toolbox → Attachment List
        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
        
        Set tabela = Nothing
        On Error Resume Next
        Set tabela = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
        On Error GoTo 0
        
        If Not tabela Is Nothing Then
            
            For i = 0 To tabela.RowCount - 1
                
                titulo = UCase(tabela.getCellValue(i, "BITM_DESCR"))
                
                If InStr(titulo, "PACK") > 0 Or InStr(titulo, "LIST") > 0 Then
                    
                    Encontrou = True
                    
                    tabela.currentCellRow = i
                    tabela.selectedRows = i
                    tabela.pressToolbarButton "%ATTA_EXPORT"
                    
                    ' Avoid overwriting existing file
                    If Dir(Caminho & tknum & ".pdf") <> "" Then
                        tknum = tknum & "_" & Format(Now, "hhmmss")
                    End If
                    
                    session.findById("wnd[2]/usr/ctxtDY_PATH").Text = Caminho
                    session.findById("wnd[2]/usr/ctxtDY_FILENAME").Text = tknum & ".pdf"
                    session.findById("wnd[2]/tbar[0]/btn[0]").press
                    
                    Exit For
                    
                End If
                
            Next i
            
        End If
        
        ' Close attachment window
        On Error Resume Next
        session.findById("wnd[1]").Close
        On Error GoTo 0
        
        ' ====================================================
        ' Mark result in column B
        ' ====================================================
        
        If Encontrou Then
            ws.Cells(Linha, 2).Value = "OK"
        Else
            ws.Cells(Linha, 2).Value = "ERRO"
        End If
        
        Linha = Linha + 1
        
    Loop
    
    MsgBox "Processo finalizado!", vbInformation

Saida:
    Set tabela = Nothing
    Set session = Nothing
    Set SAPCon = Nothing
    Set SAPApp = Nothing
    Set SapGuiAuto = Nothing
    Set ws = Nothing
    Exit Sub

ErroHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
    Resume Saida
    
End Sub
