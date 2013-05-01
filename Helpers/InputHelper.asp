<!-- #include File="../Helpers/EncodingHelper.asp" -->
<%
	Dim input : Set input = New InputHelperClass

	Class InputHelperClass
		Public Function GetParameter(name)
			GetParameter = Request.Form(encoder.encode(name))
		End Function
	End Class
%>
