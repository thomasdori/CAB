<!--#Include File = "CustomTemplate.asp" -->
<%
	Call  PageController.RenderPage()
	Public Function Page_Configure()
		PageController.PageTitle = "Admin Only"
		PageController.AuthorizedRoles =  Array("Admin")
		PageController.PageCSSFiles = "main.css" 		
	End function
%>
<%Public Function Page_RenderForm()%>
	<H2>An Admin Page</H2>
<%End Function%>
