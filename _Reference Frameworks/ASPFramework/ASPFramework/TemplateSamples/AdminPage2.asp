<!--#Include File = "CustomTemplate.asp " -->
<%
Call  PageController.RenderPage()
Public Function Page_Configure()
	PageController.PageTitle = "Mixed Roles"
	PageController.AuthorizedRoles =  Array("Admin","Manager")
	PageController.PageCSSFiles = "main.css"
End Function 
%>

<%Public Function Page_RenderForm()%>	
	<H2>Both, Admin and Manager roles can access this page!</H2>
<%End Function%>
