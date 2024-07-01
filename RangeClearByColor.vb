Option Explicit
Sub Range_Clear()
Dim LastRow As Long
Dim LastColumn As Long
Dim ws As Worksheet
Dim oList As ListObject
Dim ColorFound As Variant
Dim CopyAndPasteColumnColor As Variant
Dim StartCopyAndPasteCell As String
Dim cell As Range

With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    .DisplayStatusBar = False
    .EnableEvents = False
End With

''' Adjusting if needed
CopyAndPasteColumnColor = 40
'Application.FindFormat.Clear 'Ensure Find Formatting Rule is Reset
For Each ws In Worksheets
    ws.Activate ' Go to this sheet
    For Each oList In ws.ListObjects
        For Each cell In oList.HeaderRowRange ' Searching cells within first row
'            Set oList = ws.ListObjects(oList.Name) it might be redundant that's why I took it off
            oList.AutoFilter.ShowAllData ' Take out filter
'            Application.FindFormat.Interior.colorIndex = CopyAndPasteColumnColor ' Store active cell's fill color into "Find"
            
            If cell.Interior.colorIndex = CopyAndPasteColumnColor Then
'                ColorFound = oList.HeaderRowRange.Find("", , , , , , , , True).Address(False, False) ' Find the cell color as you want
                ColorFound = cell.Address(False, False)
                StartCopyAndPasteCell = ws.Range(ColorFound).Offset(1, 0).Address
                LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                LastColumn = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
                ws.Range(StartCopyAndPasteCell & ":" & Col_Letter(LastColumn) & LastRow).ClearContents ' Clear contents based on the color
                ws.Range(StartCopyAndPasteCell & ":" & Col_Letter(LastColumn) & LastRow).Select
                
                Debug.Print "Table name:                " & oList.Name
                Debug.Print "Color Found:               " & ColorFound
                Debug.Print "Start Copy and Paste Cell: " & StartCopyAndPasteCell
                Debug.Print "End Copy and Paste Cell:   " & Col_Letter(LastColumn) & LastRow
                Debug.Print vbNewLine
                Application.FindFormat.Clear 'Ensure Find Formatting Rule is Reset
            Else
                   'Nothing
            End If
        Next cell
    Next oList
Next ws

With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .DisplayStatusBar = True
    .EnableEvents = True
End With

End Sub


Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
