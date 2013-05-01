<%Public Function Page_Configure()
	PageController.SecuredHomeURL = "PublicPage.asp"
	PageController.PublicHomeURL  = "PublicPage.asp"
	PageController.RequiresAuthentication  = False
	PageController.PageCSSFiles = "main.css"
End Function 
%>

<%Public Function Page_RenderHeader()%>	
	<!--#Include File = "TemplateHeader.asp" -->
<%End Function
Public Function Page_RenderFooter()%>
	<!--#Include File = "TemplateFooter.asp" -->
<%End Function%>
