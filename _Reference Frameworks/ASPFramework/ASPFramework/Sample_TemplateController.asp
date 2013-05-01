<!--#Include File = "Controls\PageController.asp" -->
<%
Public Function Page_Configure()
	PageController.RequiresAuthentication  = False
	PageController.PageCSSFiles = "Samples.css"
End Function 

Public Function Page_RenderHeader()
%>
	<!--#Include File = "Home.asp" -->
	<Span Class="Caption"><%=PageController.PageTitle%> sample</Span>
<%
End Function

Public Function Page_RenderFooter()
End Function
%>



