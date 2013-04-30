<!-- #include File="../../Helpers/ErrorHelper.asp" -->
<%
	Dim footer
	Set footer = new FooterPartial

	Class FooterPartial
		Function getContent
			getContent = error.getCustomErrors & error.getStingerErrors & error.getAspErrors & "</body></html>"
		End Function
	End Class
%>