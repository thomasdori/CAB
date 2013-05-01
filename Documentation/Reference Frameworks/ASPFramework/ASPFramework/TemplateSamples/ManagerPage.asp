<!--#Include File = "CustomTemplate.asp" -->
<%
Call  PageController.RenderPage()
Public Function Page_Configure()
	PageController.PageTitle = "Secured Page"
	PageController.AuthorizedRoles =  Array("Manager")
	PageController.PageCSSFiles = "main.css"
End Function 
%>

<%Public Function Page_RenderForm()%>	
	<H2>A Manager Page!</H2>
<%End Function%>
