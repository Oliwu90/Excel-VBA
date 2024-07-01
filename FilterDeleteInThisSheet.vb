
Function FilterDeleteInThisSheet(strFilterHeaderName As String _
, strSingleMutipleValue As String _
, strThisSheetName As String _
, strThisTableName As String _
)

''''''''purpose:delete rows in a table after given filter value
''''''''   note:"strSingleMutipleValue" could be single or multiple values between quotation. for example, "apple". or "apple, orange"

    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
    End With
    ''' return column number
    Dim rngFind As Range
    Set rngFind = Worksheets(strThisSheetName).Rows(1).Find(strFilterHeaderName, LookIn:=xlValues)
    FilterIndexNumber = rngFind.Column
    
    ''' make filter array
    ''' filterOut() can't be Variant since it would give you mistypematch. need to declare as string
    Dim filterOut() As String

    ''' store input into array
    filterOut() = Split(strSingleMutipleValue, ",")

    ''' filter by input from strSingleMutipleValue
    Worksheets(strThisSheetName).ListObjects(strThisTableName).Range.AutoFilter Field:=FilterIndexNumber, Criteria1:=filterOut(), Operator:=xlFilterValues
    
    ''' "rwosAfterFilter" determind if anything finds after filter
    Dim rowsAfterFilter As Integer
    rowsAfterFilter = Worksheets(strThisSheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count

    If rowsAfterFilter = 1 Then
        ''' do nothing becasue nothing found
    Else
        ''' Delete rows if find filter value
        Dim PIVOT_TABLE As ListObject: Set PIVOT_TABLE = Worksheets(strThisSheetName).ListObjects(strThisTableName)
        PIVOT_TABLE.DataBodyRange.Rows.Delete
        
        ''' unfilter
        Worksheets(strThisSheetName).ListObjects(strThisTableName).AutoFilter.ShowAllData
    End If
    
    Worksheets(strThisSheetName).ListObjects(strThisTableName).AutoFilter.ShowAllData
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .EnableEvents = True
    End With

End Function

