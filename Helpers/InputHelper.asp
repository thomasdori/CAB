<%
	'Dependencies: EncodingHelper.asp
	Dim input : Set input = New InputHelperClass

	Class InputHelperClass
		Public Function GetString(name)
			'GetParameter = Request.Form(name)
			GetString = Request.Form(encoder.encode(name))
		End Function

		'Todo: implement GetLng, GetInt, GetDate, GetTime, GetDbl, GetEmail
	End Class
%>
