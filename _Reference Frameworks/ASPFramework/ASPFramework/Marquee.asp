<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Marquee.asp"    -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<TITLE>Server Marquee Example</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Samples.css">
	</HEAD>
	<BODY>
		<!--Include File = "Home.asp" -->
		<%	Page.Execute %>
		<Span Class="Caption">Marquee<hr></Span>
		<%Page.OpenForm%>
		<B style='font-size:10px'>Marquee Text:</b>&nbsp;<%cmdUpdate%>
		<br><%txtContents%>
		<hr>
		<%oMarquee%>
		<BR>
		<%Page.CloseForm%>
	</BODY>
</HTML>
<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim oMarquee
	Dim txtContents
	Dim cmdUpdate
	
	Public Function Page_Init()
		Set cmdUpdate   = New_ServerLinkButton("cmdUpdate")
		Set txtContents = New_ServerTextArea("txtContents",40,10)
		Set oMarquee = New_ServerMarquee("oMarquee")
		oMarquee.Control.EnableViewState = False
		oMarquee.Behavior = "alternate"
		oMarquee.Width    = "320px"
	End Function

	Public Function Page_Controls_Init()						
		txtContents.Text = "<-- CLASP Was Here -->"
		cmdUpdate.Text = "Update"
	End Function	
	
	Public Function Page_Load()
		oMarquee.Text = txtContents.Text
	End Function
	

%>
