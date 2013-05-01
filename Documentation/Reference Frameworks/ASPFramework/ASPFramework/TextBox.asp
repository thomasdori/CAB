<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<script language="javascript">
	function HandleKBPress(ctrl) {
			if(event && event.type == 'keypress' && event.keyCode==13) {
			var obj = clasp.getObject(ctrl);
			if(obj)
				obj.click();
		}
	}
</script>

</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">TEXTBOX EXAMPLE</Span>

<%Page.OpenForm%>
<table border=0>
	<tr valign=top><td><%txtFirstName%></td><td><%cmdReadOnlyYN%></td></tr>
	<tr valign=top><td><%txtLastName%></td><td><%cmdVisibleYN%></td></tr>
	<tr valign=top><td><%txtTextArea%></td><td><%cmdEnabledYN%></td></tr>
	<tr valign=top><td><%txtPassword%></td><td></td></tr>
</table>
<%'req1%>
<HR>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	'Response.AddHeader  "content-disposition", "attachment; filename=filename.ext"	
	
	Dim txtFirstName
	Dim txtLastName
	Dim txtTextArea
	Dim txtPassword
	Dim cmdReadOnlyYN
	Dim cmdVisibleYN
	Dim cmdEnabledYN
	Public Function Page_Init()
		
		Set txtFirstName = New_ServerTextBox("txtFirstName")
		Set txtLastName  = New_ServerTextBox("txtLastName")
		Set txtTextArea  = New_ServerTextBox("txtTextArea")
		Set txtPassword  = New_ServerTextBox("txtPassword")
		Set cmdReadOnlyYN = New_ServerLinkButton("cmdReadOnlyYN")
		Set cmdVisibleYN = New_ServerLinkButton("cmdVisibleYN")
		Set cmdEnabledYN = New_ServerLinkButton("cmdEnabledYN")
		'txtFirstName.AutoPostBack = True
		Page.RegisterEventListener txtFirstName,"onkeypress","HandleKBPress('cmdVisibleYN')"
		
	End Function

	Public Function Page_Controls_Init()
		txtFirstName.Caption = "First Name"
		txtLastName.Caption  = "Last Name"		
		
		txtTextArea.Caption = "Text Area"		
		txtTextArea.Mode = 3 'Text Area
		txtTextArea.Rows=4
		txtTextArea.Cols = 40
		
		txtPassword.Caption = "Password"
		txtPassword.Mode = 2 'Password
		txtPassword.MaxLength = 10 'Restrict Size
		txtPassword.Size = 10 'Set Width
		
		cmdReadOnlyYN.Text = "Toggle ReadOnly"
		cmdVisibleYN.Text  = "Toggle Visible"
		cmdEnabledYN.Text  = "Toggle Enabled"
				
		txtLastName.RaiseOnChanged = True
		
		
	End Function

	Public Function Page_Load()
		If txtFirstName.TextChanged Then
			Response.Write "<BR>First Name was modified: Detected by reading property TextChanged<BR>"
		End If
		'Page.SetDefaultObject txtTextArea,"focus"
	End Function

	Public Function txtLastName_OnChanged(obj,args)
		Response.Write "<BR>You changed the last name. Detected using OnChanged event...<BR>"
	End Function
	
	Public Function cmdReadOnlyYN_OnClick()
		txtFirstName.ReadOnly = Not txtFirstName.ReadOnly
	End Function

	Public Function cmdVisibleYN_OnClick()		
		txtLastName.Control.Visible = (Not txtLastName.Control.Visible)
	End Function
	
	Public Function cmdEnabledYN_OnClick()
		txtTextArea.Control.Enabled = Not txtTextArea.Control.Enabled
	End Function
	
%>