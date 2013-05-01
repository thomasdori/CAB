<%
	Dim output : Set output = New OutputHelperClass

	Class OutputHelperClass
		Public Function Prepare(value)
			Prepare = encoder.encode(value)
		End Sub

		Public Function PrepareUrl(value)
			PrepareUlr = encoder.encodeUrl(value)
		End Sub

		Public Sub Write(value)
			Response.Write(Prepare(value))
		End Sub

		Public Sub WriteLine(value)
			Write(value & vbCrLf)
		End Sub

		Public Sub WriteUrl(value)
			Response.Write(PrepareUrl(value))
		End Sub
	End Class
%>