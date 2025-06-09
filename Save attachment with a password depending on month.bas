Attribute VB_Name = "PW_Save"
Public Sub PW_Cobalt_Save(itm As Outlook.MailItem)
    Dim objAtt As Outlook.Attachment
    Dim saveFolder As String
    Dim AttachmentList As Attachments
    Dim OneAttachment As Attachment
    Dim OneFileType As String
    Dim AttachmentName As String
    Dim AttachmentFullName As String
    Dim dateFormat As String
    Dim strDate As String
    'the current date
    strDate = Date
    'Dim xlApp As Object
    'Dim fso As Scripting.FileSystemObject
    'Set fso = New Scripting.FileSystemObject
    Dim MyPath As String

    saveFolder = "\\cru-pro-01\ProDisk\Folder\"

   dateFormat = Format(itm.ReceivedTime, "dd-mm-yyyy-HH-mm-ss")
   For Each OneAttachment In itm.Attachments
          ' check if csv or else
        ' The code looks last 4 characters,
        ' including period and will work as long
        ' as you use 4 characters in each extension which should be the case anyway
        'MsgBox (OneAttachment.fileName)
                OneFileType = LCase$(Right$(OneAttachment.FileName, 4))
                ' Select Case File type
                Select Case OneFileType
                    Case ".csv"
                    'betting rid of the extension to only keep the file name
                    If (InStr(OneAttachment.FileName, ".") > 0) Then
                    AttachmentName = Left(OneAttachment.FileName, InStr(OneAttachment.FileName, ".") - 1)
                    End If
                'now we can get the proper name and add the extension when needed
                        AttachmentName = "FILE NAME " & dateFormat
                        AttachmentFullName = AttachmentName & ".csv"
                    ' we save the csv
                       OneAttachment.SaveAsFile saveFolder & AttachmentFullName
                       ' calling the password function
                       ' the file name is without the extension.
                       ProtectExcelWorkbook saveFolder & AttachmentName
          Set objAtt = Nothing

End Select
Next
End Sub

Sub ProtectExcelWorkbook(filePath As String)
    Dim dateFormat As String
    Dim MyDate As String
    MyDate = Date
    Dim pwd As String
    Dim saveName As String
    Dim xlsApp As Excel.Application
    Dim xlsFile As Excel.Workbook
    'the current date
    dateFormat = Format(MyDate, "mm-yy")
        If dateFormat = "09-21" Then pwd = "Pinea"
        If dateFormat = "10-21" Then pwd = "Choco"
        If dateFormat = "11-21" Then pwd = "Chall"
        If dateFormat = "12-21" Then pwd = "Enter"
        If dateFormat = "01-22" Then pwd = "Affect"
        If dateFormat = "02-22" Then pwd = "Import&"
        If dateFormat = "03-22" Then pwd = "Pack"
        If dateFormat = "04-22" Then pwd = "Curio"
        If dateFormat = "05-22" Then pwd = "Deli"
        If dateFormat = "06-22" Then pwd = "Chemi"
        If dateFormat = "07-22" Then pwd = "Magnes"
        If dateFormat = "08-22" Then pwd = "Gene"
        If dateFormat = "09-22" Then pwd = "Radi"
        If dateFormat = "10-22" Then pwd = "Sun"
        If dateFormat = "11-22" Then pwd = "Fire"
        If dateFormat = "12-22" Then pwd = "Emp"
        If dateFormat = "01-23" Then pwd = "Water"
        If dateFormat = "02-23" Then pwd = "Sig"
        If dateFormat = "03-23" Then pwd = "Forgo"
        If dateFormat = "04-23" Then pwd = "Young"
        If dateFormat = "05-23" Then pwd = "Tinker"
        If dateFormat = "06-23" Then pwd = "Furnit"
        If dateFormat = "07-23" Then pwd = "Snow"
        If dateFormat = "08-23" Then pwd = "Sent"
        If dateFormat = "09-23" Then pwd = "Metro"
    'MsgBox (filePath)
    'pwd = "Pinea"
    Set xlsApp = New Excel.Application
    ' we add the ext csv to read the csv
    xlsApp.Workbooks.Open (filePath & ".csv")
    ' we add the ext xlsx to save the file
    '
'Manipulate Excel file here
On Error Resume Next
'MsgBox "Start xlx.Insert Shift"
    xlsApp.ActiveWorkbook.ActiveSheet.Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'MsgBox "Start xlx.D1 FormulaR1C1 Accident"
xlsApp.ActiveWorkbook.ActiveSheet.Range("D1").FormulaR1C1 = "Accident Dte"
'MsgBox "Start xlx. D2 FormulaR1C1 Date"
xlsApp.ActiveWorkbook.ActiveSheet.Range("D2").FormulaR1C1 = _
        "=IFERROR(DATEVALUE(TEXT(RC[-1],""mm/dd/yyyy"")),RC[-1])"
'MsgBox "Start xlx.AutoFill Destination"
xlsApp.ActiveWorkbook.ActiveSheet.Range("D2").AutoFill Destination:=Range("D2:D" & Range("C" & Rows.Count).End(xlUp).Row)
'MsgBox "Start xlx, Select"
xlsApp.ActiveWorkbook.ActiveSheet.Columns("D:D").Select
'MsgBox "End xlx"
    '
    xlsApp.ActiveWorkbook.ActiveSheet.Columns("D:D").Copy
    xlsApp.ActiveWorkbook.ActiveSheet.Columns("D:D").PasteSpecial Paste:=xlPasteValues
    xlsApp.ActiveWorkbook.ActiveSheet.Columns("C:C").Delete Shift:=xlToLeft
    xlsApp.ActiveWorkbook.ActiveSheet.Range("A1").Select
    xlsApp.ActiveWorkbook.SaveAs FileName:=filePath & ".xlsx", FileFormat:=51, Password:=pwd
           
    Set xlsFile = Nothing
    xlsApp.Quit
    Set xlsApp = Nothing
Kill filePath & ".csv"
End Sub

