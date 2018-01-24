'Get file name in a folder and paste into the worksheet

Sub GetSubFilesFileNames()

Range("C6:AZ6").ClearContents

    Dim DirectoryPath As String
    Dim Filename As String
        
    DirectoryPath = Range("C5") & "\"
    Filename = Dir(DirectoryPath)
    
    Range("B6").Select
    
    Do While Filename <> ""
    
        ActiveCell.Offset(0, 1).Activate
        ActiveCell = Filename
       Filename = Dir
       
    Loop

    Range("C6", Range("C6").End(xlToRight)).Select
    Selection.Replace What:=".xlsm", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False

    Range("C6", Range("C6").End(xlToRight)).Select
    Selection.Replace What:=".xlsx", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False
    
    Range("B6").Select
    
    MsgBox "All file names in sub folder were copied over"
End Sub