<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Timer.asp"    -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Timer Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">

</HEAD>
<BODY id="CLASPBody">

<!--Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">Timer<hr></Span>
<center><h2><%lblOutput%></h2></center>
<%tmrTimer%><HR> 	<!--I don't have to be in the Form because I'm not an html input control! -->
<%Page.OpenForm%>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim tmrTimer
	Dim lblOutPut
	
	
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		
		Set lblOutput	= New_ServerLabel("lblOutput")
		Set tmrTimer	= New_ServerTimer("tmrTimer",1000) 'defaults to 1000 millseconds
		
		tmrTimer.Interval = 3000
		
	End Function

	Function Page_Controls_Init()	
		lblOutput.Text = "Last Date & Time:" & Now
	End Function

	Function tmrTimer_OnTimer()
		lblOutput.Text =  "Last Date & Time:" & Now
	End Function

%>