Attribute VB_Name = "ImportaSAP"
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object
Public ThisWb As Workbook
Public wsMF As Worksheet
Public wsO As Worksheet
Public wsC As Worksheet
Public tbl1 As ListObject
Public tbl2 As ListObject
Public tbl3 As ListObject
    
Sub ImportarMateraisSAP(Optional HideFromMacroList = True)

    Dim response As VbMsgBoxResult
    
    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo Terminate  ' Exit the function or sub
        End If
    Loop
    
    Set ThisWb = ThisWorkbook
    
    On Error Resume Next
    Set wsMF = ThisWorkbook.Sheets("Materiais Faltantes")
    Set wsO = ThisWorkbook.Sheets("Obras")
    Set wsC = ThisWorkbook.Sheets("Contatos")
    On Error GoTo 0
    
    If wsMF Is Nothing Or wsO Is Nothing Then
        MsgBox "Planilhas não encontradas.", vbInformation
        GoTo Terminate
    End If
    
    Set tbl1 = wsMF.ListObjects("Tabela1")
    Set tbl2 = wsO.ListObjects("Tabela2")
    Set tbl3 = wsC.ListObjects("Tabela3")
    
    ExportSAPData
    
    EndSAPScripting
    
    AddMateriaisToLista
    
    EnviaEmail
    
Terminate:
    
    ' Desliga a otimização
    OptimizeCodeExecution False

End Sub

Function EnviaEmail()
    Dim newWb As Workbook
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim emailCell As Range
    Dim emailList As String
    Dim newHTMLBody As String
    Dim filePath As String
    Dim fileNumber As Integer
    
    If Dir(ThisWb.Path & "\email_base.html") = "" Then
        MsgBox "O e-mail base não foi encontrado.", vbExclamation
        Exit Function
    End If
    
    ' Create an instance of Outlook
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject(Class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    If outlookApp Is Nothing Then
        MsgBox "O Outlook não está instalado nesse computador.", vbExclamation
        Exit Function
    End If
    
    ' Build the email list from the table
    emailList = ""
    For Each emailCell In tbl3.ListColumns("Contatos").DataBodyRange ' Change "Email" to the column name with email addresses
        If emailCell.Value <> "" Then
            emailList = emailList & emailCell.Value & "; "
        End If
    Next emailCell
    
    ' Remove the trailing semicolon and space
    If Len(emailList) > 2 Then
        emailList = Left(emailList, Len(emailList) - 2)
    End If
    
    ' Prepare the attachment
    Set newWb = Workbooks.Add
    
    ' Copy the first and second worksheets from the current workbook to the new workbook
    wsMF.Copy After:=newWb.Worksheets(newWb.Worksheets.Count)
    wsO.Copy After:=newWb.Worksheets(newWb.Worksheets.Count)
    newWb.Worksheets(1).Delete
    
    newWb.Worksheets("Materiais Faltantes").ListObjects("Tabela1").DataBodyRange.Cells(1, 1).Formula = "=VLOOKUP([@Ordem],Tabela2,2,FALSE)"
    newWb.Worksheets("Materiais Faltantes").ListObjects("Tabela1").DataBodyRange.Cells(1, 2).Formula = "=VLOOKUP([@Ordem],Tabela2,3,FALSE)"
    
    newWb.SaveAs ThisWb.Path & "\Lista de Materiais Faltantes.xlsx"
    
    ' Create a new email
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Path to the HTML file
    filePath = ThisWorkbook.Path & "\email_base.html"
    
    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Insert the image before "Obrigada"
    newHTMLBody = Input$(LOF(fileNumber), fileNumber)
    Close fileNumber
    
    ' Replace the old image reference with the new path
    newHTMLBody = Replace(newHTMLBody, "email_base_arquivos/image002.png", ThisWorkbook.Path & "\email_base_arquivos\image002.png")
    newHTMLBody = Replace(newHTMLBody, "email_base_arquivos/image001.png", ThisWorkbook.Path & "\email_base_arquivos\image001.png")
    
    ' Set the email properties
    With outlookMail
        .To = emailList
        .Subject = "Lista de Materiais Faltantes"
        .BodyFormat = 2 ' 2 = olFormatHTML
        
        ' Body content and user signature
        .HTMLBody = newHTMLBody
        .Attachments.Add newWb.FullName
        ' Display the email for review before sending
        .Display
    End With
    
    ' Clean up
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    ' Close the workbook without saving changes
    newWb.Close SaveChanges:=False
    
    'Application.Wait (Now + TimeValue("00:00:03"))
    
    ' Delete the workbook file
    On Error Resume Next ' In case the file is not found or cannot be deleted
    Kill ThisWb.Path & "\Lista de Materiais Faltantes.xlsx"
    On Error GoTo 0
    
End Function

Function AddMateriaisToLista()
    
    Dim SAPwb As Workbook
    Dim wb As Workbook
    Dim SAPSheet As Worksheet
    Dim Row As Range
    Dim targetName As String
    Dim ordem As String
    Dim found As Boolean
    
    ' Name of the workbook to find
    targetName = "export.XLSX"
    found = False
    
    ' Loop through all open workbooks
    For Each wb In Application.Workbooks
        If UCase(wb.Name) = UCase(targetName) Then
            Set SAPwb = wb
            found = True
            Exit For
        End If
    Next wb
    
    If Not found Then
        Set SAPwb = Workbooks.Open(ThisWb.Path & "\" & targetName)
    End If
    
    Set SAPSheet = SAPwb.Sheets(1)
    
    SAPSheet.Rows(1).Delete
    
    For Each Row In SAPSheet.Rows
        With Row
            If .Cells(1, 1) = "" And .Cells(1, 3) = "" Then
                Exit For
            ElseIf .Cells(1, 3) = "" Then
                ordem = .Cells(1, 1).Value
            ElseIf .Cells(1, 1) = "" Then
                .Cells(1, 1) = ordem
            End If
        End With
    Next Row
    
    For Each Row In SAPSheet.Rows
        With Row
            If .Cells(1, 1) = "" And .Cells(1, 2) = "" Then
                Exit For
            ElseIf .Cells(1, 3) = "" Then
                .Delete
            End If
        End With
    Next Row
    
    tbl1.DataBodyRange.ClearContents
    
    tbl1.Resize wsMF.Range("A1:P" & SAPSheet.Cells(1, 1).End(xlDown).Row + 1)
    
    SAPSheet.Cells(1, 1).CurrentRegion.Copy
    
    tbl1.DataBodyRange.Cells(1, 3).PasteSpecial Paste:=xlPasteAll
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    tbl1.DataBodyRange.Cells(1, 1).Formula = "=VLOOKUP([@Ordem],Tabela2,2,FALSE)"
    tbl1.DataBodyRange.Cells(1, 2).Formula = "=VLOOKUP([@Ordem],Tabela2,3,FALSE)"
     
    ' Close the workbook without saving changes
    SAPwb.Close SaveChanges:=False
    
    'Application.Wait (Now + TimeValue("00:00:03"))
    
    ' Delete the workbook file
    On Error Resume Next ' In case the file is not found or cannot be deleted
    Kill ThisWb.Path & "\" & targetName
    On Error GoTo 0
    
End Function

Function ExportSAPData()

    ' Copia as DRs da Tabela 2
    tbl2.ListColumns(1).DataBodyRange.Copy

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZTPP092"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "HENCKE"
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    Dim grid As Object
    Dim iRow As Long
    Dim searchValue As String
    Dim colName As String
    
    ' Set your search value and the column name (as defined in the grid)
    searchValue = "FALTANTES_PLA"
    colName = "VARIANT"   ' Replace with the actual column name
    
    ' Get the grid control
    Set grid = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    
    ' Loop through all rows
    For iRow = 0 To grid.RowCount - 1
        If grid.GetCellValue(iRow, colName) = searchValue Then
            ' When found, set the current cell to the matching row
            grid.CurrentCellRow = iRow
            ' Depending on your setup, SelectedRows may require a string
            grid.SelectedRows = CStr(iRow)
            ' Double-click the cell to perform the action
            grid.doubleClickCurrentCell
            Exit For   ' Exit the loop once the desired row is found
        End If
    Next iRow

    session.findById("wnd[0]/usr/btn%_S_NETWK_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = "01.01.2018"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[46]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ThisWb.Path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "export.XLSX"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    Application.CutCopyMode = False
        
End Function

Function SetupSAPScripting() As Boolean
    
    Dim isHomePage As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    Set Connection = SAPApplication.Children(0)
    Set session = Connection.Children(0)
    
    SetupSAPScripting = True
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function
