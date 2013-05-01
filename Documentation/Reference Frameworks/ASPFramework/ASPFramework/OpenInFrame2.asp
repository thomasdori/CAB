<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<TITLE>Page.OpenInFrame Example</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Samples.css">
	</HEAD>
	<BODY id="CLASPBody">
		<%Page.Execute%>
		<Span Class="Caption">Page.OpenInFrame Example Page #2</Span>
		<HR>
		<%lblMessage%>
		<HR>
		<%Page.OpenForm%>
			<%cmdFrameA%> | <%cmdFrameB%>
		<%Page.CloseForm%>
	</BODY>
</HTML>
<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim lblMessage
	Dim cmdFrameA
	Dim cmdFrameB


'PAGE EVENT HANDLERS
	Function Page_Init()
		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdFrameA = New_ServerButton("cmdFrameA")
		Set cmdFrameB = New_ServerButton("cmdFrameB")

		Page.ViewStateMode = VIEW_STATE_MODE_CLIENT

	End Function

	Function Page_Controls_Init()
		cmdFrameA.Text = "Load This Page (Page 2) in Top Frame"
		cmdFrameB.Text = "Load Page 1 in Top Frame"
	End Function

	Function Page_Load()
	End Function

'WEBCONTROLS POSTBACK EVENT HANDLERS

	Function cmdFrameA_OnClick()
		lblMessage.Text = "Page 2 is now in Top Frame"
		Page.OpenInFrame "topframe","" ' by omitting url the current page will be loaded into the select frame
	End Function

	Function cmdFrameB_OnClick()
		lblMessage.Text = "Page 1 is now in Top Frame"
		Page.OpenInFrame "topframe","OpenInFrame1.asp"
	End Function

%>
