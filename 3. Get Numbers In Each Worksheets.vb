Sub RetrieveNumbers()

Dim NumberFiles As Integer, FilesVisited As Integer, RowNumber As Integer
Let NumberFiles = ActiveSheet.Cells("2", "A").Value
Let FilesVisited = 0                            'start from 0
Let RowNumber = 4                            'start from column B
If NumberFiles > 30 Then
    MsgBox "Don't try to retrieve numbers from more than 30 files at a time!"
Else
    For FilesVisited = 1 To NumberFiles
        
        'Open files, get path, file, tab name and cells
        Dim PathFileOpen As String, NameFileOpen As String, NameTab As String, FileDir As String
        Let PathFileOpen = ActiveSheet.Cells(RowNumber, "A").Text
        Let NameFileOpen = ActiveSheet.Cells(RowNumber, "B").Text
        Let NameTab = ActiveSheet.Cells(RowNumber, "C").Text
        
        Dim N As Integer, Cell As String, NumberYears As String, FullLink As String
        NumberYears = ActiveSheet.Cells("2", "B").Value
        For N = 4 To NumberYears + 3
            Cell = ActiveSheet.Cells(RowNumber, N).Text
            FullLink = "(=)'" & PathFileOpen & "\[" & NameFileOpen & ".xlsm]" & NameTab & "'!" & Cell
            ActiveSheet.Cells(RowNumber, N + 13).Value = FullLink
        Next N
        RowNumber = RowNumber + 1
    Next FilesVisited
End If

ActiveSheet.Range("A1").CurrentRegion.Replace What:="(=)", Replacement:="=", _
LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

MsgBox ("Finished Retrieving")

End Sub


Private Sub CommandButton21_Click()
    RetrieveNumbers
End Sub