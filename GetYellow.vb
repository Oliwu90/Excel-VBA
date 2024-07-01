Option Explicit
Sub GetYellow()

With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
End With


Dim ThisWb As Workbook: Set ThisWb = ActiveWorkbook
Dim ThisWks As Worksheet
Dim oList As ListObject
Dim wsPaste As Worksheet


For Each ThisWks In ThisWb.Worksheets
    For Each oList In ThisWks.ListObjects
        oList.AutoFilter.ShowAllData
    Next oList
Next ThisWks

Dim n As Integer
Dim colorIndex As Integer 'yellow = 6
Dim xPath As String

Dim FileExtentionFormat As String
Dim GetBook As String

FileExtentionFormat = ThisWb.FileFormat

''' find normal workbook or macro workbook
Select Case FileExtentionFormat
    Case xlWorkbookDefault '51 xlsx
        GetBook = Replace(ThisWb.Name, ".xlsx", "")
    Case xlOpenXMLWorkbookMacroEnabled '52 xlsm
        GetBook = Replace(ThisWb.Name, ".xlsm", "")
    Case xlExcel12 '50 xlsb
        GetBook = Replace(ThisWb.Name, ".xlsb", "")
End Select

Dim sFileName As String
sFileName = GetBook & " ,V"

xPath = ActiveWorkbook.Path & "\"

Dim Deswb As Workbook: Set Deswb = Workbooks.Add

For n = 1 To ThisWb.Sheets.Count
    colorIndex = ThisWb.Sheets(n).Tab.colorIndex
        If colorIndex = 6 Then
            ThisWb.Sheets(n).Copy after:=Deswb.Sheets(Deswb.Sheets.Count)
        Else
         'DO NOTHING
        End If
Next n

''For Each wsPaste In Deswb.Sheets
    ''wsPaste.Cells.Copy
    ''wsPaste.Range("A1").PasteSpecial Paste:=xlValues
''Next wsPaste

Deswb.Sheets(1).Delete
Deswb.Sheets(1).Activate
Deswb.SaveAs Filename:=xPath & sFileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook

With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .DisplayStatusBar = True
    .EnableEvents = True
End With

End Sub

