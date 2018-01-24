Sub Copy and Pasting between two workbooks()

'Find how many files need to be created
Dim NumberFiles As Integer
Dim FilesVisited As Integer
Dim ColumnNumber As Integer

Let NumberFiles = Worksheets("File_List").Cells("4", "C").Value
Let FilesVisited = 0                            'start from 0
Let ColumnNumber = 3                            'start from column C

If NumberFiles > 50 Then
    MsgBox "Don't try to visit more than 50 files at a time!"
Else
    Dim SourcePathFileOpen As String
    Dim SourceNameFileOpen As String
    Dim SourceFileDir As String
    
    Let SourcePathFileOpen = Worksheets("File_List").Cells(1, 3).Text
    Let SourceNameFileOpen = Worksheets("File_List").Cells(2, 3).Text
    Let SourceFileDir = SourcePathFileOpen & "\" & SourceNameFileOpen
    
    Workbooks.Open Filename:=SourceFileDir, UpdateLinks:=0
    
    For FilesVisited = 1 To NumberFiles
    'start from the first file to be created. There are too many rows below, so from next line of code, indentation will restart.
        
        Application.Calculation = xlCalculationManual

        'Open file and unprotect it
        Dim PathFileOpen As String
        Dim NameFileOpen As String
        Dim FileDir As String
        Dim TabName As String
        Dim N As Integer
        Dim TabCreated As Integer
        
        Let PathFileOpen = Worksheets("File_List").Cells("5", ColumnNumber).Text
        Let NameFileOpen = Worksheets("File_List").Cells("6", ColumnNumber).Text
        Let FileDir = PathFileOpen & "\" & NameFileOpen
        
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
        
        Workbooks(SourceNameFileOpen).Activate
        
        Set Sourcetemp = ActiveWorkbook
        Application.DisplayAlerts = True
        Application.AskToUpdateLinks = True
        
        Workbooks.Open Filename:=FileDir, UpdateLinks:=0
        Workbooks(NameFileOpen).Activate
        
        ColumnNumber = ColumnNumber + 1
        
        Set Desttemp = ActiveWorkbook
        
        'SAMPLE'
        
        Sourcetemp.Worksheets("Assumptions").Range("G302:Q343").Copy
        Desttemp.Worksheets("Assumptions").Range("G302:Q343").PasteSpecial Paste:=xlPasteValues
            'Service Company Cost Allocation
            'Copy as well as the subtotal
        
        Sourcetemp.Worksheets("Assumptions").Range("H530:DW570").Copy
        Desttemp.Worksheets("Assumptions").Range("H530:DW570").PasteSpecial Paste:=xlPasteValues
            'LTD Maturities - EXTERNAL DEBT
            'not copying the subtotal
        
        Sourcetemp.Worksheets("Assumptions").Range("H575:Q615").Copy
        Desttemp.Worksheets("Assumptions").Range("H575:Q615").PasteSpecial Paste:=xlPasteValues
            'LTD Interest Expense- EXTERNAL DEBT
            'not copying the subtotal
            
        Sourcetemp.Worksheets("Assumptions").Range("H620:DW660").Copy
        Desttemp.Worksheets("Assumptions").Range("H620:DW660").PasteSpecial Paste:=xlPasteValues
            'LTD Maturities - INTERCOMPANY
            'not copying the subtotal
        
        Sourcetemp.Worksheets("Assumptions").Range("H665:Q705").Copy
        Desttemp.Worksheets("Assumptions").Range("H665:Q705").PasteSpecial Paste:=xlPasteValues
            'LTD Interest Expense- INTERCOMPANY
            'not copying the subtotal
        Sourcetemp.Worksheets("Assumptions").Range("H1080:Q1120").Copy
        Desttemp.Worksheets("Assumptions").Range("H1080:Q1120").PasteSpecial Paste:=xlPasteValues
            'CapEx spend
      
        'close file
        Desttemp.Save
        Desttemp.Close
        Application.DisplayAlerts = True
        Application.CutCopyMode = False

        Application.Calculation = xlCalculationAutomatic
    
    Next FilesVisited
    
    'Close master file
    Sourcetemp.Close False
End If

MsgBox "In total " & FilesVisited - 1 & " sub files have been visited and fixed."


End Sub