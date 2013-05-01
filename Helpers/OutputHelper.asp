<%
	Dim output : Set output = New OutputHelperClass

	Class OutputHelperClass
		Public Sub Write(value)
			Response.Write(Server.HTMLEncode(value))
		End Sub

		Public Sub WriteLine(value)
			Write(value & vbCrLf)
		End Sub

		Public Sub WriteURL(value)
			Response.Write(Server.URLEncode(value))
		End Sub
	End Class
%>