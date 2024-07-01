Sub SaveAsOnlyValues()
Dim ws As Worksheet

With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
End With

For Each ws In Worksheets
    ws.Cells.Copy
    ws.Cells.PasteSpecial xlPasteValuesAndNumberFormats
Next ws

With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .DisplayStatusBar = True
    .EnableEvents = True
    .CutCopyMode = False
End With

End Sub
