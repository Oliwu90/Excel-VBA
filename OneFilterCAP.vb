
Function OneFilterCAP(strFilterHeaderName_1 As String _
, strFilterValue_1 As String _
, strCopyHeaderName As String _
, strCopySheetName As String _
, strCopyTableName As String _
, strPastedColumnName As String _
, strPasteSheetName As String _
, strPasteTableName As String _
, defXlWholeOrXlPart As String)
'***
' one filter in sheet specified and copy and paste to another sheet specified
'***
    ''' somehow needs to activate the sheet you want to examine if it's fine to CAP
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(strCopySheetName)
    ws.Activate
    
    ''' make an array to stor multiple filter values, and sperate by comma
    strFilterArray_1 = Split(strFilterValue_1, ",")

    ''' filter by header
    Dim rngFind As Range
    
    '''start to filter
    With ws.ListObjects(strCopyTableName)
       
        Set rngFind = Worksheets(strCopySheetName).Rows(1).Find(strFilterHeaderName_1, LookIn:=xlValues)
        FilterIndexNumber = rngFind.Column
        .Range.AutoFilter Field:=FilterIndexNumber, Criteria1:=strFilterArray_1, Operator:=xlFilterValues

    End With
    
    ''' rwosAfterFilter determind if anything finds after filter
    Dim rowsAfterFilter As Integer
    rowsAfterFilter = Worksheets(strCopySheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count

    If rowsAfterFilter = 1 Then
        ''' do nothing becasue nothing found
    Else
        CountFilterRows = Worksheets(strCopySheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
        ''' resize table
        Dim PIVOT_TABLE As ListObject: Set PIVOT_TABLE = Worksheets(strPasteSheetName).ListObjects(strPasteTableName)
        PIVOT_TABLE.Resize PIVOT_TABLE.Range.Resize(CountFilterRows)
        Worksheets(strPasteSheetName).Range("A" & CountFilterRows + 1 & ":XFD1048576").Clear
        
        ''' 1st copy
        Set rngFind = Worksheets(strCopySheetName).Rows(1).Find(strCopyHeaderName, LookIn:=xlValues)
        CopyHeaderIndexNumber = rngFind.Column
        Worksheets(strCopySheetName).ListObjects(strCopyTableName).ListColumns(CopyHeaderIndexNumber).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        
        ''' paste to a new work book
        Dim objOpenNewWorkbook As Workbook
        Set objOpenNewWorkbook = Workbooks.Add
        objOpenNewWorkbook.Sheets(1).Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        
        ''' 2nd: copy
        objOpenNewWorkbook.Sheets(1).UsedRange.Copy
        
        ''' need to go back to template
        ThisWorkbook.Activate
        
        Dim rngPasteFind As Range: Set rngPasteFind = Worksheets(strPasteSheetName).Rows(1).Find(strPastedColumnName, LookIn:=xlValues, LookAt:=defXlWholeOrXlPart)
        strPastedHeader = rngPasteFind.Address
        
        ''' 2nd paste: to template
        Worksheets(strPasteSheetName).Range(strPastedHeader).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        
        ''' close the new workbook
        objOpenNewWorkbook.Close Savechanges:=False
        
        Worksheets(strCopySheetName).ListObjects(strCopyTableName).AutoFilter.ShowAllData
    End If
    Worksheets(strCopySheetName).ListObjects(strCopyTableName).AutoFilter.ShowAllData
End Function


