<!--#Include File = "CustomTemplate.asp" -->
<%
Call  PageController.RenderPage()

Public Function Page_Configure()
	PageController.PageTitle = "Public Page"
	PageController.RequiresAuthentication  = False
	PageController.PageCSSFiles = "main.css"
End Function 
%>

<%Public Function Page_RenderForm()%>
	
	<SPAN><H2>WELCOME</H2></SPAN><BR>
	This is a public page. Click on any of the links above or to your right to access secured pages.
	You will see how you will be forced to login or you will get a no access page.
<%End Function%>
