Sub GetNumbers()
Dim PathFileOpen As String, NameFileOpen As String, FullFileName As String, PastedWorkBook As String

Let PathFileOpen = Worksheets("Summary").Cells("9", "B").Text
Let NameFileOpen = Worksheets("Summary").Cells("10", "B").Text
Let FullFileName = PathFileOpen & "\" & NameFileOpen & ".xlsm"
 
PastedWorkBook = ActiveWorkbook.Name

If InStr(PastedWorkBook, ".") > 0 Then
   PastedWorkBook = Left(PastedWorkBook, InStr(PastedWorkBook, ".") - 1)
End If
        
Application.Calculation = xlCalculationManual
Application.AskToUpdateLinks = False
Workbooks.Open Filename:=FullFileName
Application.ScreenUpdating = False

    Workbooks(NameFileOpen).Worksheets("Base_Inputs").Cells("28", "D").Resize(1, 10).Copy         'Total revenue in Residential sales
    Workbooks(PastedWorkBook).Worksheets("Summary").Cells("16", "D").Resize(1, 10).PasteSpecial xlPasteValues
        
    Workbooks(NameFileOpen).Worksheets("Base_Inputs").Cells("132", "D").Resize(1, 10).Copy        'Total Other Water Revenue in Miscellaneous
    Workbooks(PastedWorkBook).Worksheets("Summary").Cells("20", "D").Resize(1, 10).PasteSpecial xlPasteValues
        
    Workbooks(NameFileOpen).Worksheets("Base_Inputs").Cells("17", "D").Resize(1, 10).Copy       'Average # of Customers
    Workbooks(PastedWorkBook).Worksheets("Summary").Cells("26", "D").Resize(1, 10).PasteSpecial xlPasteValues
                
Application.CutCopyMode = False
Workbooks(NameFileOpen).Close
ActiveWorkbook.Save
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub