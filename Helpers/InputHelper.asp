<%
	'Dependencies: EncodingHelper.asp
	Dim input : Set input = New InputHelperClass

	Class InputHelperClass
		Public Function GetParameter(name)
			'GetParameter = Request.Form(name)
			GetParameter = Request.Form(encoder.encode(name))
		End Function

		Public Function GetParameter(name, type)
			'Todo: cast
			GetParameter = GetParameter(name)
		End Function
	End Class
%>
