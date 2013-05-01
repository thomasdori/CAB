<%
	Class List
		Dim list()

		Public Sub Add(value)
			Dim newSize : Set newSize = Ubound(list) + 1
		    ReDim Preserve list(newSize)
		    list(newSize) = value
		End Sub

		Public Sub ForEach(fnc)
			Dim i

			For i = 0 To Ubound(list)
				fnc(list(i))
			Next
		End Sub
	End Class
%>