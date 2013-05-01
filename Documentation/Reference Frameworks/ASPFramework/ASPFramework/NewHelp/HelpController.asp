<!--#Include File = "../Controls/PageController.asp" -->
<%
Dim mLeftSectionTitle 
Dim mRightSectionTitle
Public Function Page_Configure()
	'viewstate has to be on client side so it does not collide/compete with session state
	Page.ViewStateMode = VIEW_STATE_MODE_CLIENT
	Page.AutoResetScrollPosition = False
	PageController.RequiresAuthentication  = False
	PageController.PageCSSFiles = "help.css"	
	PageController.PageTitle = "CLASP 2.0"	
End Function 
%>

<%Public Function Page_RenderHeader()%>	
	<!--#Include File = "header.asp" -->
<%End Function
Public Function Page_RenderFooter()%>
	<!--#Include File = "footer.asp" -->
<%End Function%>
