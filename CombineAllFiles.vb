Option Explicit

Public Sub CombineAllFiles(strFileNamePattern As String _
, strSheetNameDes As String _
, strColNameDes As String _
, strTableNameDes As String)

'~~> below becomes parts of the function arguments
'    Dim strFileNamePattern As String
'    strFileNamePattern = "pattern_xxx_xxxx_xxxx"
'    Dim strSheetNameDes As String
'    strSheetNameDes = "Sheet1"
    
    Dim strDesPath As String
    strDesPath = Application.ActiveWorkbook.Path
    
    Dim strSourceFile As String
    
    Dim strSourceFiles As String
    strSourceFiles = strDesPath & "\" & strFileNamePattern & "*.xls*"
    
    Dim wbDes As Workbook: Set wbDes = ThisWorkbook 'if you want to consolidate files in this workbook
    Dim wsDes As Worksheet: Set wsDes = wbDes.Sheets(strSheetNameDes) 'replace Sheet1 to suit
    Dim tbDes As ListObject: Set tbDes = wsDes.ListObjects(strTableNameDes)
    
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    
    Dim longLastRowDes As Long
    Dim longLastRowSourceFile As Long
    
    With Application
            .DisplayAlerts = False
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .EnableEvents = False
    End With
    
    strSourceFile = Dir(strSourceFiles)

    Do While strSourceFile <> ""
        Debug.Print Len(strSourceFile)
        '~~> Open the file and at the same time, set your variable
        Set wbSource = Workbooks.Open(Filename:=strDesPath & "\" & strSourceFile, ReadOnly:=True)
        Set wsSource = wbSource.Sheets(1) 'I used index, you said there is only 1 sheet
        
        With wsDes
            '~~> Find column from destination file
            '~~> below becomes parts of the function arguments
            '            Dim strColNameDes As String
            '            strColNameDes = "Case Number"
            Dim rngFindDes As Range: Set rngFindDes = wsDes.Cells.Find(strColNameDes, LookIn:=xlValues)
            Dim strDesHeader As String
            strDesHeader = Split(rngFindDes.Address, "$")(1) '~~> address by alphabet
            
            Dim longColNumberDes As Long
            longColNumberDes = rngFindDes.Column '~~> address by number
            
            '~~> Find column from source file(s)
            Dim strSourceFileColumn As String
            strSourceFileColumn = strColNameDes
            Dim rngFindSource As Range: Set rngFindSource = wsSource.Cells.Find(strSourceFileColumn, LookIn:=xlValues)
            Dim strSourceFileHeader As String
            strSourceFileHeader = Split(rngFindSource.Address, "$")(1)

            '~~> Find the last rows from destination and source files
            longLastRowDes = tbDes.ListColumns(1).Range.Rows.Count '<-- last row in Column A in your Table
            longLastRowSourceFile = wsSource.Range(strSourceFileHeader & .Rows.Count).End(xlUp).Row  'dynamic last row from source files

            '~~> copy and paste juicy part
            wsSource.Range(rngFindSource.Address, Split(rngFindSource.End(xlToRight).Address, "$")(1) & longLastRowSourceFile).Copy
            tbDes.Range(longLastRowDes, strDesHeader).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
            Application.CutCopyMode = False
        End With
        
        '~~> Close the opened file
        wbSource.Close False 'set to false, because we opened it as read-only
        Set wsSource = Nothing
        Set wbSource = Nothing
        
        '~~> Load the new file
        strSourceFile = Dir()
    Loop

'~~> best way: becasue I know it's going to have blanks row
'~~>           and always copy and paste with the header from source file
    With tbDes
            '~~>remove blanks row
            .ListColumns(strColNameDes).DataBodyRange.SpecialCells(xlCellTypeBlanks).Rows.Delete
            
            '~~>remove starting header row based on header name
            .Range.AutoFilter Field:=longColNumberDes, Criteria1:=strColNameDes
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            .AutoFilter.ShowAllData
    End With

'~~> alternative way: not good becasue of loop
'    With wsDes
'        longLastRowSourceFile = wsDes.Range(strSourceFileHeader & .Rows.Count).End(xlUp).Row
'        Dim i As Long
'        '~~> always loop backwards when deleting rows
'        For i = longLastRowSourceFile To 2 Step -1
'            If .Range(strDesHeader & i).Value2 = strColNameDes Or IsEmpty(.Range(strDesHeader & i)) Then
'                .Rows(i).EntireRow.Delete
'            End If
'        Next i
'    End With

    wbDes.Save
'    wbDes.SaveAs _
'        Filename:=(strDesPath & "\" & strFileNamePattern & "_Finished Consolidated_" & Format(Now(), "mmddyy") & ".xlsx"), _
'        FileFormat:=xlOpenXMLWorkbook
    wsDes.Range("A1").Select
    
    With Application
            .DisplayAlerts = True
            .ScreenUpdating = True
            .DisplayStatusBar = True
            .EnableEvents = True
    End With

End Sub


