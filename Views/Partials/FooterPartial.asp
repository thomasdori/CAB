<!-- #include File="Helpers/ErrorHelper.asp" -->
<%
	Dim footer
	Set footer = new FooterPartial

	Class FooterView
		Function getContent
			getContent = 	Call error.getCustomErrors &
			  					Call error.getStingerErrors &
								Call error.getAspErrors &
								"</body></html>"
		End Function
	Endl Class
%>