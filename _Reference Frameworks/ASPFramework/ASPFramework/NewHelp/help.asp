<!--#Include File = "../Controls/WebControl.asp"-->
<!--#Include File = "../Controls/Server_Button.asp" -->
<!--#Include File = "../Controls/Server_TextBox.asp" -->
<!--#Include File = "../Controls/Server_RichTextBox.asp" -->
<!--#Include File = "../Controls/Server_Label.asp" -->
<!--#Include File = "SendMail.asp" -->
<%Page.Execute%>
<html>
	<head>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<title>Support</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<LINK rel="stylesheet" type="text/css" href="help.css">
	</head>
	<body>
<%Page.OpenForm%>
	<Span style="font-weight:bold;font-size:16px">Classic ASP Framework</Span><hr>
	<%If Not Page.IsPostBack Then
		If Request.QueryString("done")<>"" Then
			Response.Write "<span style='font-weight:bold;color:maroon'>Your support request was sent, you will be contacted shortly!</span><br>"
		End If
	  End If
	%>
	<table class="download_section" ID="Table1"><tr><td>&nbsp;Help Request</td></tr></table>		
		<table cellpadding=0 cellspacing=0 border=0 ID="Table4">
			<tr>
				<td>Email:&nbsp;&nbsp;</td>
				<td><%txtFrom%></td><td valign=top style='color:red'>&nbsp;</td>
			</tr>
			<tr>
				<td>Subject&nbsp;&nbsp;</td>
				<td><%txtSubject%></td><td valign=middle style='color:red'>&nbsp;</td>
			</tr>
			<tr>
				<td valign=top>Body&nbsp;&nbsp;</td>
				<td><%txtBody%></td><td valign=middle style='color:red'>&nbsp;</td>
			</tr>
			<tr><td></td>
				<td colspan="2"  style="color:maroon"><%cmdSubmit%>&nbsp;<%lblMsg%></td>
			</tr>				
		</table>
	</table>
	<hr>	
<%Page.CloseForm%>
	</body>
</html>


<%
Dim txtFrom
Dim txtSubject
Dim txtBody
Dim cmdSubmit
Dim lblMsg

Public Function Page_Init()
	Set txtFrom   = New_ServerTextBoxEx("txtFrom",50,30)
	Set txtSubject = New_ServerTextBoxEx("txtSubject",50,30)
	Set txtBody   = New_ServerRichTextBox("txtBody")
	Set cmdSubmit = New_ServerButton("cmdSubmit")
	Set lblMsg      = New_ServerLabel("lblMsg")
End Function

Public Function Page_Controls_Init()
	cmdSubmit.Text = "Submit"
End Function

Public Function Page_Load()
	lblMsg.Text = ""
End Function

Public Function cmdSubmit_OnClick()
	If SendSupportMail(txtFrom.Text,IIF(txtSubject.Text="","SUPPORT",txtSubject.Text),txtBody.Text) Then
		Response.Redirect "help.asp?done=1"
	End If
End Function


%>