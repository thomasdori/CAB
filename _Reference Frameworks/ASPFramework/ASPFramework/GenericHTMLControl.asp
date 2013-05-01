<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_HTMLGenericControl.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">Generic HTML Control (using div tag)<hr></Span>
<%Page.OpenForm%>
<table border=0>
	<tr valign=top><td valign=top>Html</td><td valign=top><%txtHtml%></td></tr>
	<tr valign=top><td colspan=2><%htmlGeneric%></td></tr>
</table>
	<%cmdShow%> | <%cmdDoNothing%>
<HR>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim txtHtml
	Dim htmlGeneric
	Dim cmdShow	
	Dim cmdDoNothing
	
	Public Function Page_Init()
		Set txtHtml = New_ServerTextBox("txtHTML")
		Set htmlGeneric  = New_ServerHTMLGenericControl("htmlGeneric","DIV")
		Set cmdShow  = New_ServerLinkButton("cmdShow")
		Set cmdDoNothing = New_ServerLinkButton("cmdDoNothing")
	End Function

	Public Function Page_Controls_Init()
		
		txtHtml.Mode = 3 'Text Area
		txtHtml.Rows=10
		txtHtml.Cols = 60		
		cmdShow.Text = "Show Html"
		cmdDoNothing.Text = "Do nothing"
	End Function

	Public Function Page_Load()
	End Function

	
	Public Function cmdShow_OnClick()
		htmlGeneric.InnerHTML = txtHtml.Text		
	End Function

%>