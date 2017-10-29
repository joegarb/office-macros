REM Concatenate the text from a range of cells to a single cell with newline delimiter
REM Use as a formula in a cell, eg.=CONCATRANGE(A2:C5)

Function ConcatRange(ByVal range As Variant)
	If Not IsArray(range) Then
		Exit Function
	End If
	
	Dim result As String, current As String
	
	For i = lbound(range, 1) To ubound(range, 1)
		For j = lbound(range, 2) To ubound(range, 2)
			current = range(i, j)
			
			If current <> "" And current <> 0 Then
				If result <> "" Then
					result = result & Chr(10) 'newline
				End If
				
				result = result & current
			End If
		Next j
	Next i
	
	ConcatRange = result
End Function
