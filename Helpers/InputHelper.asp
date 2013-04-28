<%
Dim input
set input = new InputHelper

Class InputHelper
	Function getParameter(name)
		getParameter = Request.Form(name)
	End Function
End Class
%>
