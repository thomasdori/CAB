<%
	'Dependcies: EncodingHelper.asp

	Dim input
	set input = new InputHelper

	Class InputHelper
		Function getParameter(name)
			getParameter = Request.Form(encoder.encode(name))
		End Function
	End Class
%>
