<!--#Include File = "CustomTemplate.asp" -->
<%
Call  PageController.RenderPage()
Public Function Page_Configure()
	PageController.PageTitle = "No Access"
	PageController.RequiresAuthentication  = True
	PageController.RequiresAuthorization   = False
	PageController.PageCSSFiles = "main.css"
End Function 
%>

<%Public Function Page_RenderForm()%>	
	<H2>NO ACCESS!!!!!</H2>
	<BR>WHAT ARE YOU TRYING TO ACCESS EH!, I GOT YOU!!!
	<BR>
<%End Function%>
