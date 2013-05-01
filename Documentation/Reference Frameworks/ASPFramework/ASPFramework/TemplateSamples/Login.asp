<!--#Include File = "CustomTemplate.asp" -->
<!--#Include File = "..\Controls\Server_Button.asp" -->
<!--#Include File = "..\Controls\Server_TextBox.asp"    -->
<!--#Include File = "..\Controls\Server_CheckBoxList.asp"    -->
<!--#Include File = "..\Controls\Server_Label.asp"    -->
<%
Dim txtUserName
Dim txtPassword
Dim cmdLogin
Dim chkRoles
Dim lblErrors
Call  PageController.RenderPage()

Public Function Page_Configure()
	PageController.SecuredHomeURL = "PublicPage.asp"
	PageController.PublicHomeURL  = "PublicPage.asp"
	PageController.PageTitle = "Login"
	PageController.RequiresAuthentication  = False
	PageController.PageCSSFiles = "main.css"
	PageController.BodyAttributes = " marginheight=0 marginwidth=0 topmargin=0 leftmargin=0 "
End Function 
%>

<%Public Function Page_RenderForm()%>	
	<BR><BR><BR><BR><BR>
	<TABLE align="center">
		<TR><td>User Name</td><td><%txtUserName%></td></TR>
		<TR><td>Password</td><td><%txtPassword%></td></TR>
		<TR><td valign=top>Roles</td><td><%chkRoles%></td></TR>		
		<TR><td align=right colspan=2><Hr><%cmdLogin%></td></TR>		
	</TABLE>
	<%lblErrors%>

<%End Function%>


<%
':: PAGE EVENTS
Public Function Page_Init()
	Set txtUserName = new_serverTextBox("txtUserName")
	Set txtPassword = new_serverTextBox("txtPassword")
	Set chkRoles = new_serverCheckBoxList("chkRoles")
	Set cmdLogin = new_serverButton("cmdLogin")	
	Set lblErrors = new_serverLabel("lblErrors")
	lblErrors.Control.EnableViewState = False
End Function

Public Function Page_Controls_Init()
	cmdLogin.Text = "Login"
	chkRoles.Items.Append "Admin","Admin",False
	chkRoles.Items.Append "Manager","Manager",False
End Function

':: POSTBACK HANDLERS

Public Function cmdLogin_OnClick()
	Dim strErr
	Dim roles
	Set strErr = new StringBuilder
	If chkRoles.Items.GetSelectedText = "" Then
		strErr.Append  "Please select at lease one role<BR>"
	End If
	
	If txtUserName.Text = "" Or txtPassword.Text = "" Then
		strErr.Append   "User Name and Password are required fields."
	End If	
	
	lblErrors.Text = strErr.ToString()
	
	If lblErrors.Text = "" Then
		'Quick and dirty... just for sample purposes :-?
		'I addes the setroles just for the sample because Auhenticate should do it.
		
		If chkRoles.Items.IsSelected(0) Then
			Redim roles(0)
			roles(0) = "Admin"
		End If
		If chkRoles.Items.IsSelected(1) Then
			If IsArray(roles) Then
				Redim Preserve roles(1)
			Else
				Redim roles(0)
			End If
			roles(UBOUND(roles)) = "Manager"
		End If
		
	CurrentUser.RedirectFromLoginPage txtUserName.Text,"TEST","TEST","TEST@AA.COM",Roles
		
	End If	
	
End Function

%>