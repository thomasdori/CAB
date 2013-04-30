<%
	'Dependcies: EncodingHelper.asp

	Dim input
	set input = new InputHelper

	Class InputHelper
		Public Function GetParameter(name)
			GetParameter = Request.Form(encoder.encode(name))
		End Function
	End Class
%>
