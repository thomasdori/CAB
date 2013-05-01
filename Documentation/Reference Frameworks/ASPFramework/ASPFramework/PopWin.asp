<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_FieldValidator.asp"    -->

<HTML>
	<HEAD>
		<title>POPUP WINDOWS</title>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<LINK rel="stylesheet" type="text/css" href="Samples.css">	
	</HEAD>
	<BODY id="bodyid">
<!--Include File = "Home.asp"        -->		
<% Page.Execute%>
<Span Class="Caption">Popup Windows from Server<br></Span>
<%Page.OpenForm%>
	<%cmdOpen%> | <%cmdOpenC%> | <%cmdOpenL%>
	<hr>
	<span>Demonstrates the use of the Page.OpenWindow and Page.GetWindowScript functions</span>
<%Page.CloseForm%>

<%  

	Dim cmdOpenC
	Dim cmdOpen
	Dim cmdOpenL

	Public Function Page_Init()
		Set cmdOpen = New_ServerLinkButton("cmdOpen")
		Set cmdOpenC = New_ServerLinkButton("cmdOpenC")
		Set cmdOpenL = New_ServerLinkButton("cmdOpenL")

		'opens and centers a window 300x200 when the cmdOpen button is clicked
		Page.RegisterEventListener cmdOpenC,"onclick",Page.GetOpenWindowScript("Wn","datagrid.asp",-1,-1,300,200,true,true,"")

		'opens and vertically centers a window 300x200 when the cmdOpen button is clicked
		Page.RegisterEventListener cmdOpen,"onclick",Page.GetOpenWindowScript("Wn","datagrid.asp",10,-1,300,200,true,true,"")

	End Function

	Public Function Page_Controls_Init()		
		cmdOpen.Text	= "Open a Window"
		cmdOpenC.Text	= "Open a centered Window"
		cmdOpenL.Text	= "Open a Window on Page Load"
	End Function


	Public Function cmdOpenL_OnClick()		
		'opens a window 200x200 when the page loads
		Call Page.OpenWindow("xp","datagrid.asp",100,100,200,200,true,true,"")
	End Function



%>
	</BODY>
</HTML>
