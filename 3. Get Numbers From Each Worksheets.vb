Sub RetrieveNumbers(wrkSht As Worksheet)

Dim NumberFiles As Integer, FilesVisited As Integer, RowNumber As Integer
Dim PathFileOpen As String, NameFileOpen As String, NameTab As String, FileDir As String
Dim N As Integer, Cell As String, NumberYears As String, FullLink As String

Application.ScreenUpdating = False

    
NumberFiles = wrkSht.Cells("2", "A").Value
FilesVisited = 0                         'start from 0
RowNumber = 4                            'start from column B

If NumberFiles > 30 Then
        MsgBox "Don't try to retrieve numbers from more than 30 files at a time!"
    Else
        For FilesVisited = 1 To NumberFiles
            
            'Open files, get path, file, tab name and cells
            PathFileOpen = wrkSht.Cells(RowNumber, "A").Text
            NameFileOpen = wrkSht.Cells(RowNumber, "B").Text
            NameTab = wrkSht.Cells(RowNumber, "C").Text

            NumberYears = wrkSht.Cells("2", "B").Value
            For N = 4 To NumberYears + 3
                Cell = wrkSht.Cells(RowNumber, N).Text
                FullLink = "(=)'" & PathFileOpen & "\[" & NameFileOpen & ".xlsm]" & NameTab & "'!" & Cell
                wrkSht.Cells(RowNumber, N + 13).Value = FullLink
        Next N
            RowNumber = RowNumber + 1
    Next FilesVisited
End If

wrkSht.Range("A1").CurrentRegion.Replace What:="(=)", Replacement:="=", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False

Application.ScreenUpdating = True


End Sub