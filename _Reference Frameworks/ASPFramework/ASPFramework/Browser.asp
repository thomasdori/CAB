<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Browser Agent Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">

</HEAD>
<BODY id="CLASPBody">

<!--#Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">BROWSER EXAMPLE</Span>
<HR>
Version: <%lblVersion%><HR> 
AppName: <%lblAppName%><HR> 
UserAgent: <%lblUserAgent%><HR> 	<!--I don't have to be in the Form because I'm not an html input control! -->
<b>SCREEN:</b><hr style="border:1px dotted blue">
Screen Width: <%=Browser.ScreenWidth%><HR> 
Screen Height: <%=Browser.ScreenHeight%><HR> 
Screen Color Depth: <%=Browser.ScreenColor%> (bits)<HR> 
<b>BROWSER:</b><hr style="border:1px dotted blue">
IsNS: <%=Browser.IsNS%><HR> 
IsDOM: <%=Browser.IsDOM%><HR> 
IsMSIE: <%=Browser.IsMSIE%><HR> 
IsOpera: <%=Browser.IsOpera%><HR> 
IsSafari: <%=Browser.IsSafari%><HR> 
IsGecko/Moz: <%=Browser.IsGecko%><HR> 
<b>PLATFORM:</b><hr style="border:1px dotted blue">
IsMac: <%=Browser.IsMac%><HR> 
IsWin32: <%=Browser.IsWin32%><HR> 
IsUnix: <%=Browser.IsUnix%><HR> 
IsUnix: <%=Browser.IsUnix%><HR> 
<b>OTHERS:</b><hr style="border:1px dotted blue">
HasNoScript: <%=Browser.HasNoScript%><HR> 
HasNoIFRAME: <%=Browser.HasNoIFRAME%><HR> 

<%Page.OpenForm%>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim lblUserAgent
	Dim lblVersion
	Dim lblAppName
	
	
'PAGE EVENT HANDLERS	
	Function Page_Init()		
	 	Set lblUserAgent = New_ServerLabel("lblUserAgent")
		Set lblVersion = New_ServerLabel("lblVersion")
		Set lblAppName = New_ServerLabel("lblAppName")

		lblVersion.Text = Browser.Version		
		lblAppName.Text = Browser.AppName		
		lblUserAgent.Text = Browser.UserAgent
		
	End Function

	Function Page_Controls_Init()	
	End Function

	Function Page_Load()
	End Function

%>