<%
	Dim output
	Set output = new OutputHelper

	Class OutputHelper
		Sub write(value)
			Response.Write(Server.HTMLEncode(value))
		End Sub

		Sub writeURL(value)
			Response.Write(Server.URLEncode(value))
		End Sub
	End Class
%>