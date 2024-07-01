Public Sub CopyAndPasteNoFilterAdvanced(strSrcFileNameWithExtention As String _
, strSrcSheetName _
, strPastedTabName As String _
, strPastedPivotTableName As String _
, strPastedColumnName As String)

''' this function has source sheet name
Speed_Turn_On

Dim wkMasterPath As String
wkMasterPath = ActiveWorkbook.Path & "\"

Dim wkMaster As Workbook: Set wkMaster = ActiveWorkbook

Dim wsPasted As Worksheet: Set wsPasted = wkMaster.Sheets("" & strPastedTabName & "")

Dim wkSrc As Workbook: Set wkSrc = Workbooks.Open(Filename:=wkMasterPath & strSrcFileNameWithExtention)

Dim wsSrc As Worksheet: Set wsSrc = wkSrc.Sheets(strSrcSheetName)

Dim lRowwkSrc As Long
lRowwkSrc = wsSrc.Range("A" & wsSrc.Rows.Count).End(xlUp).Row 'get the last row

Dim PIVOT_TABLE As ListObject: Set PIVOT_TABLE = wsPasted.ListObjects(strPastedPivotTableName)
PIVOT_TABLE.Resize PIVOT_TABLE.Range.Resize(lRowwkSrc)
wsPasted.Range("A" & lRowwkSrc + 1 & ":XFD1048576").Clear

Dim rngFind As Range: Set rngFind = wsPasted.Rows(1).Find(strPastedColumnName, LookIn:=xlValues)
Dim strPastedHeader As String
strPastedHeader = rngFind.Address

wsSrc.UsedRange.Copy
wsPasted.Range(strPastedHeader).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
Application.CutCopyMode = False
wsPasted.Rows(2).Delete
wkSrc.Close
wsPasted.Range("A1").Select

Debug.Print "Source file                  : " & wkMasterPath & strSrcFileNameWithExtention
Debug.Print "Last row in the source file  : " & lRowwkSrc
Debug.Print "Current active workbook      : " & ActiveWorkbook.Name
Debug.Print "Paste sheet name             : " & wsPasted.Name
Debug.Print "Paste column header          : " & strPastedColumnName
Debug.Print "Paste cell                   : " & Range(strPastedHeader).Offset(1).Address
Debug.Print vbNewLine

Speed_Turn_Off

End Sub




