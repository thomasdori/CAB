<!--#Include File = "../Controls/WebControl.asp"-->
<!--#Include File = "../Controls/Server_Button.asp" -->
<!--#Include File = "../Controls/Server_TextBox.asp" -->
<!--#Include File = "../Controls/Server_FieldValidator.asp" -->
<!--#Include File = "../Controls/Server_Label.asp" -->
<!--#Include File = "../DBWrapper.asp" -->
<!--#Include File = "SendMail.asp" -->
<%Page.Execute%>
<html>
	<head>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<title>Member Registration</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<LINK rel="stylesheet" type="text/css" href="help.css">
	</head>
	<body>
<%Page.OpenForm%>
	<Span style="font-weight:bold;font-size:16px">Classic ASP Framework</Span><hr>
	<table class="download_section" ID="Table1"><tr><td>&nbsp;Member Login</td></tr></table>		
		<table cellpadding=0 cellspacing=0 border=0 ID="Table4">
			<tr>
				<td>Email:&nbsp;&nbsp;</td>
				<td><%txtLoginEmail%></td><td valign=top style='color:red'>&nbsp;<%oVLoginmail%></td>
			</tr>
			<tr>
				<td>Password:&nbsp;&nbsp;</td>
				<td><%txtLoginPwd%></td><td valign=middle style='color:red'>&nbsp;<%oVLoginPassword%></td>
			</tr>
			<tr>
				<td colspan="3"  style="color:maroon"><%cmdLogin%>&nbsp;<%lblLoginMsg%></td>
			</tr>
			<tr><td>&nbsp;</td><td colspan=2><span>forgot your password?<a href="javascript:clasp.form.doPostBack('RememberMe','Page')"> click here </a></span></td></tr>
		</table>
	</table>
	<hr>	
	<table class="download_section" ID="Table2"><tr><td>&nbsp;New User?</td></tr></table>		
		<table cellpadding=0 cellspacing=0 border=0 ID="Table3">
			<tr>
				<td>First Name:&nbsp;&nbsp;</td>
				<td><%txtFirstName%></td><td valign=top style='color:red'>&nbsp;<%oVFirstName%></td>
			</tr>
			<tr>
				<td>Last Name:&nbsp;&nbsp;</td>
				<td><%txtLastName%></td><td valign=middle style='color:red'>&nbsp;<%oVLastName%></td>
			</tr>
			<tr>
				<td>Email:</td>
				<td><%txtEmail%></td><td valign=middle style='color:red'>&nbsp;<%oVEmail%></td>
			</tr>
			<tr>
				<td>Password:</td>
				<td><%txtPassword%></td><td valign=middle style='color:red'>&nbsp;<%oVPassword%></td>
			</tr>
				<tr>
					<td>Confirm Password:</td>
					<td><%txtPasswordRe%></td><td valign=middle style='color:red'>&nbsp;<%oVPassword2%></td>
				</tr>
			<tr>
				<td colspan="3"  style="color:maroon"><%cmdSubmit%>&nbsp;<%lblRegisterMsg%></td>
			</tr>				
		</table>
<%Page.CloseForm%>
	</body>
</html>


<%
Dim txtLoginEmail
Dim txtLoginPwd
Dim oVLoginmail,oVLoginPassword
Dim cmdLogin
Dim lblLoginMsg

Dim txtFirstName
Dim txtLastName
Dim txtEmail
Dim txtPassword,txtPasswordRe
Dim cmdSubmit
Dim oVFirstName,oVLastName,oVEmail,oVPassword,oVPassword2
Dim lblRegisterMsg

Public Function Page_Init()
	Set txtLoginEmail = New_ServerTextBoxEx("txtLoginEmail",50,20)
	Set txtLoginPwd   = New_ServerTextBoxEx("txtLoginPwd",50,20)
	Set oVLoginmail   = New_ServerRequiredFieldValidator("oVLoginmail",txtLoginEmail,"* Required","","L")
	Set oVLoginPassword  = New_ServerRequiredFieldValidator("oVLoginPassword",txtLoginPwd,"* Required","","L")
	Set cmdLogin = New_ServerButton("cmdLogin")
	Set lblLoginMsg      = New_ServerLabel("lblLoginMsg")
	cmdLogin.ValidationGroup = "L"
	
	Set txtFirstName = New_ServerTextBoxEx("txtFirstName",50,20)
	Set txtLastName = New_ServerTextBoxEx("txtLastName",50,20)
	Set txtEmail = New_ServerTextBoxEx("txtEmail",100,30)
	Set txtPassword = New_ServerTextBoxEx("txtPassword",10,10)
	Set txtPasswordRe = New_ServerTextBoxEx("txtPasswordRe",10,10)

	Set cmdSubmit = New_ServerButton("cmdSubmit")
	Set oVFirstName = New_ServerRequiredFieldValidator("oVFirstName",txtFirstName,"* Required","","")
	Set oVLastName  = New_ServerRequiredFieldValidator("oVLastName",txtLastName,"* Required","","")
	Set oVEmail	    = New_ServerRequiredFieldValidator("oVEmail",txtEmail,"* Required","","")
	Set oVPassword  = New_ServerRequiredFieldValidator("oVPassword",txtPassword,"* Required","","")
	Set oVPassword2  = New_ServerRequiredFieldValidator("oVPassword2",txtPasswordRe,"* Required","","")

	Set lblRegisterMsg = New_ServerLabel("lblRegisterMsg")

End Function

Public Function Page_Controls_Init()
	cmdSubmit.Text = "Submit"
	cmdLogin.Text = "Login"
	txtPassword.Mode = 2
	txtLoginPwd.Mode = 2
End Function

Public Function Page_Load()
	lblRegisterMsg.Text = ""	
	lblLoginMsg.Text = ""
End Function

Public Function cmdSubmit_OnClick()
	Dim sSQL
	Dim rs
	txtPassword.Text = Trim(txtPassword.Text)
	txtPasswordRe.Text = Trim(txtPasswordRe.Text)
	If Len(txtPassword.Text) < 6 Then
		lblRegisterMsg.Text  = "Passwords need to be at least 6 characters"
		txtPassword.Text = ""
		txtPasswordRe.Text = ""
		Exit Function
	End If
	If txtPassword.Text <>  txtPasswordRe.Text  Then
		lblRegisterMsg.Text  = "Password and Confirm Password don't match"
		txtPassword.Text = ""
		txtPasswordRe.Text = ""
		Exit Function
	End If

	sSQL = "SELECT Count(*) From CLASPUser WHERE Email='" & Replace(txtEmail.Text,"'","''") & "'"
	Set rs = DbLayer.GetRecordSet(sSQL)
	If rs(0).Value = 0 Then
		sSQL = "Insert Into CLASPUser (FirstName,LastName,Email,Password,IsAdmin) VALUES('"  &_
			Replace(txtFirstName.Text,"'","''") & "','" & _
			Replace(txtLastName.Text,"'","''") & "','" & _
			Replace(txtEmail.Text,"'","''") & "','" & _
			Replace(txtPassword.Text,"'","''") & "',0)"
				
		DbLayer.ExecuteSQL sSQL
		Response.Cookies.Item("CLASP_SiteUser") = txtFirstName.Text  & " " & txtLastName.Text
		Response.Cookies.Item("CLASP_SiteEmail") = txtEmail.Text
		'SendMail "webmaster@claspdev.com","CLASP Account Created",txtFirstName.Text  & " " & txtLastName.Text
		Page.RegisterClientScript "XXXX","<script language='javascript'>top.clasp.form.doPostBack('OnAccCreated','Page')</script>"';
	Else
		lblRegisterMsg.Text =  "A user with the same email is already registered..."
	End If
	txtPassword.Text = ""
End Function

Public Function Page_RememberMe(e)
	Dim sSQL
	Dim rs

	If txtLoginEmail.Text = "" Then
		lblLoginMsg.Text = "At least enter the email!!!, Jezzz!"
		Exit Function
	End If
	
	sSQL = "SELECT * FROM CLASPUser WHERE Email = '" & Replace(txtLoginEmail.Text,"'","''") & "'"
	Set rs = DBLayer.GetRecordSet(sSQL)
	If rs.RecordCount = 0 Then
		lblLoginMsg.Text = "Email address not found in our databases..."
		Exit Function
	Else
		StartSession rs("FirstName").value,rs("LastName").value,rs("Email").value

	End If
	
End Function

Public Function cmdLogin_OnClick()
	Dim sSQL
	Dim rs
	sSQL = "SELECT * FROM CLASPUser WHERE Email = '" & Replace(txtLoginEmail.Text,"'","''") & "' AND Password='" & Replace(txtLoginPwd.Text,"'","''") & "'"
	Set rs = DBLayer.GetRecordSet(sSQL)
	If rs.RecordCount = 0 Then
		lblLoginMsg.Text = "User Not Found"
	Else
		SendSupportMail txtEmail.Text,"RESET - PWD","Please reset this user's password"
		StartSession rs("FirstName").value,rs("LastName").value,rs("Email").value
	End If
	txtLoginPwd.Text = ""
End Function

Public Function StartSession(FirstName,LastName,Email)
	Response.Cookies.Item("CLASP_SiteUser") = FirstName  & " " & LastName
	Response.Cookies.Item("CLASP_SiteEmail") = Email
	Page.RegisterClientScript "XXXX","<script language='javascript'>top.clasp.form.doPostBack('OnAccCreated','Page','1')</script>"';
End Function
%>