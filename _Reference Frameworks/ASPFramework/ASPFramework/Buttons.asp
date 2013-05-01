<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<TITLE>Server Buttons Example</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Samples.css">
	</HEAD>
	<BODY id="CLASPBody">
		<%Page.Execute%>
		<Span Class="Caption">Buttons Sample<hr></Span>
		<br><%lblMessage%>
		<%Page.OpenForm%>
			<%cmdLinkButton%> | <%cmdButton%> | <%cmdAdvanceButton%> | <%cmdImageButton%>
		<%Page.CloseForm%>
	</BODY>
</HTML>
<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim lblMessage
	Dim cmdLinkButton
	Dim cmdAdvanceButton	
	Dim cmdButton
	Dim cmdImageButton	
	
	
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdLinkButton = New_ServerLinkButton("cmdLinkButton")
		Set cmdButton = New_ServerButton("cmdButton")
		Set cmdAdvanceButton = New_ServerAdvanceButton("cmdAdvanceButton")		
		Set cmdImageButton = New_ServerImageButton("cmdImageButton")	
	End Function

	Function Page_Controls_Init()	
		cmdButton.Text = "Normal Button"
		cmdLinkButton.Text = "Link Button"
		cmdAdvanceButton.Text = "<b><font color='red'>A</font><font color='maroon'>dvance</font></b> <font color='navy'>Button</font>"
		cmdImageButton.Image = "images/book.gif"
		cmdImageButton.RollOverImage = "images/clear_all.gif"
		lblMessage.Control.Style = "border:1px solid blue;background-color:#AAAAAA"
	End Function

	Function Page_Load()
	End Function

'WEBCONTROLS POSTBACK EVENT HANDLERS	
	Function cmdButton_OnClick()
		lblMessage.Text = "You Clicked a normal button"		
	End Function

	Function cmdAdvanceButton_OnClick()
		lblMessage.Text = "You Clicked an Advance button"
	End Function

	Function cmdLinkButton_OnClick()
		lblMessage.Text = "You Clicked  Link Button"
	End Function
	
	Function cmdImageButton_OnClick()
		lblMessage.Text = "You Clicked on a Image Button"
	End Function

'SUPPORTING FUNCTIONS (IF ANY)
%>
