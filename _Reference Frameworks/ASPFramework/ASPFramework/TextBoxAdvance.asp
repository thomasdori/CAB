<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->

<!-- By Raymond Irving (xwisdom) -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Formatted Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Page.Execute()%>	
<Span Class="Caption">Formatted TextBoxes</Span>
<Br>By Raymond Irving<br>
<%Page.OpenForm%>
<table border=0>
	<tr valign=top><td><%txtNormal%></td></tr>
	<tr valign=top><td><%txtUpper%></td></tr>
	<tr valign=top><td><%txtLower%></td></tr>
	<tr valign=top><td><%txtNumeric%></td></tr>
	<tr valign=top><td><%txtDate%></td></tr>
	<tr valign=top><td><%txtDate2%></td></tr>
	<tr valign=top><td><%txtDate3%></td></tr>
</table>
<HR>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	'Response.AddHeader  "content-disposition", "attachment; filename=filename.ext"	
	
	Dim txtNormal
	Dim txtUpper
	Dim txtLower
	Dim txtNumeric
	Dim txtDate
	Dim txtDate2
	Dim txtDate3
	
	
	'Public Function Page_HandleCacheRequest()
	'	Dim obj
	'	
	'	'Only attempt to get from cache if not trying to cache it ! (?Cache=true)
	'	If Request.QueryString("Cache")="" Then		
	'		Set obj = Server.CreateObject("CLASP.PageCache")
	'		Call obj.GetCachedResponse("TextBox.asp")
	'		Response.End
	'	End If
	''	Application.Contents("")
	'	
	'End Function
	
	Public Function Page_Init()
		
		Set txtNormal 	= New_ServerTextBox("txtNormal")
		Set txtUpper  	= New_ServerTextBox("txtUpper")
		Set txtLower  	= New_ServerTextBox("txtLower")
		Set txtNumeric  = New_ServerTextBox("txtNumeric")
		Set txtDate 	= New_ServerDateTextBox("txtDate")
		Set txtDate2 	= New_ServerDateTextBox("txtDate2")
		Set txtDate3 	= New_ServerDateTextBox("txtDate3")

	End Function

	Public Function Page_Controls_Init()
		txtNormal.Caption = "^Normal"
		txtNormal.Mode = 1
		
		txtUpper.Caption  = "^Upper"		
		txtUpper.Mode = 5
		
		txtLower.Caption = "^Lower"
		txtLower.Mode = 6
		
		txtNumeric.Caption = "^Numeric"
		txtNumeric.Mode = 7
		txtNumeric.Text = "1,080.00"
		txtNumeric.FormatString = "#,##0.00"

		txtDate.Caption = "^Date"
		txtDate.Text = "5/4/2004"		
	End Function

	Public Function Page_Load()
		If txtNormal.TextChanged Then
			Response.Write "<BR>Data was modified<BR>"
		End If
	End Function


%>