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
		<title>CLASP V2.0 Downloads</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<LINK rel="stylesheet" type="text/css" href="help.css">
	</head>
	<body>
		<Span style="font-weight:bold;font-size:16px">Classic ASP Framework</Span><hr>
		<table class="download_section"><tr><td>&nbsp;Download</td></tr></table>		
		You can download the code <a style="color:navy;font-weight:bold" href="http://www.gotdotnet.com/workspaces/releases/checkfordownload.aspx?id=69b08b15-d456-4cf9-8b12-d4642ef0c22e&ReleaseId=6825ca1b-29a3-4ed9-af49-e9d493b4a87b">here</a><br>
		<!--You can download the release notes <a style="color:navy;font-weight:bold" href="#">here</a><br>-->
		
		<hr>
<%Page.OpenForm%>
<%If Not  IsAuthenticated Then%>
		<table class="download_section" ID="Table1"><tr><td>&nbsp;Register Now: Register to receive the latest news and fixes. A special section will open up and you will be able to download the newst controls and patches before they are released!!</td></tr></table>		
			<table cellpadding=0 cellspacing=0 border=0>
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
			<hr>
<%Else%>
	<a href="grid.zip"><b>New DataGrid!</b></a> Raymond has been working on upgrades such as automatic row highlighting, improved sort styles and much more.</a><br>
	<br><a href="ASPFramework122804.zip"><b>Last build as of  12-28-04!</b></a></a><br>
	<br><a href="ClaspAjax.zip"><b>Ajax Update files, check it out!</b></a></a><br>

<%End If%>
<%Page.CloseForm%>
<!--
		<table class="download_section" ID="Table2"><tr><td>&nbsp;Donate...</td></tr></table>		
			<IMG SRC="/ASPFramework/images/paypal_logo.gif">
			<BR>Donate:
			If you think CLASP is helpful and want to donate please do so at: AAA
			<br>
		The contributions will be used to maintain this website (and hopefully setup a 
		better one with more space with dotnetnuked as the frontend!
-->
	</body>
</html>


<%
Dim txtFirstName
Dim txtLastName
Dim txtEmail
Dim txtPassword,txtPasswordRe
Dim cmdSubmit
Dim oVFirstName,oVLastName,oVEmail,oVPassword,oVPassword2
Dim lblRegisterMsg
Dim IsAuthenticated

Public Function Page_Init()
	Set txtFirstName = New_ServerTextBoxEx("txtFirstName",50,20)
	Set txtLastName = New_ServerTextBoxEx("txtLastName",50,20)
	Set txtEmail = New_ServerTextBoxEx("txtEmail",100,30)
	Set txtPassword   = New_ServerTextBoxEx("txtPassword",10,10)
	Set txtPasswordRe = New_ServerTextBoxEx("txtPasswordRe",10,10)
	Set cmdSubmit = New_ServerButton("cmdSubmit")
	Set oVFirstName = New_ServerRequiredFieldValidator("oVFirstName",txtFirstName,"* Required","","")
	Set oVLastName  = New_ServerRequiredFieldValidator("oVLastName",txtLastName,"* Required","","")
	Set oVEmail	    = New_ServerRequiredFieldValidator("oVEmail",txtEmail,"* Required","","")
	Set oVPassword  = New_ServerRequiredFieldValidator("oVPassword",txtPassword,"* Required","","")
	Set oVPassword2  = New_ServerRequiredFieldValidator("oVPassword2",txtPasswordRe,"* Required","","")
	
	Set lblRegisterMsg = New_ServerLabel("lblRegisterMsg")
	IsAuthenticated = Request.Cookies("CLASP_SiteUser") <> ""
End Function

Public Function Page_Controls_Init()
	cmdSubmit.Text = "Submit"
	txtPassword.Mode = 2
	txtPasswordRe.Mode = 2
End Function

Public Function Page_Load()
	lblRegisterMsg.Text = ""	
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
		'SendMail txtFirstName.Text & " " & txtLastName.Text, txtEmail.Text,"CLASP Account Created","Thanks for registering with CLASP!"
		'SendMail "webmaster@claspdev.com","CLASP Account Created",txtFirstName.Text  & " " & txtLastName.Text
		Page.RegisterClientScript "XXXX","<script language='javascript'>top.clasp.form.doPostBack('OnAccCreated','Page','0')</script>"';
	Else
		lblRegisterMsg.Text =  "A user with the same email is already registered..."
	End If
	txtPassword.Text = ""
End Function
%>