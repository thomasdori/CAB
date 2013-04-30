<%
	'Dependcies: EncodingHelper.asp

	Dim input
	set input = New InputHelperClass

	Class InputHelperClass
		Public Function GetParameter(name)
			GetParameter = Request.Form(encoder.encode(name))
		End Function
	End Class
%>
