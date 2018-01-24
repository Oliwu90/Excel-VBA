         'Find and replace inter-workbook links
	        ThisWorkbook.Activate
	        
	        Dim replacename As String
	        Dim MasterFileNameType As Boolean
	        replacename = Range("F2").Value
	        MasterFileNameType = Range("G2").Value
	  
	        Desttemp.Worksheets(1).Activate
	        
	            If MasterFileNameType = True Then
	                For Each Worksheet In ActiveWorkbook.Worksheets
	                    Worksheet.Activate
	                    Cells.Select
	                    Selection.Replace replacename, "'", xlPart, xlByRows, False 'execute the replacement with parameters.
	                Next Worksheet
	

	            Else
	                For Each Worksheet In ActiveWorkbook.Worksheets
 	                    Worksheet.Activate
 	                    Cells.Select
 	                    Cells.Replace what:=replacename, Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False 'execute the replacement with parameters.
 	                Next Worksheet
 	            End If

		End Sub 
