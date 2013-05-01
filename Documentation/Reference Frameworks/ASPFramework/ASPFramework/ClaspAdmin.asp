<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_CheckBox.asp"     -->
<!--#Include File = "Controls\Server_Label.asp"     -->
<!--#Include File = "Controls\Server_Timer.asp"    -->
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<TITLE>Basic Administration Page</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Samples.css">
	</script>
	</HEAD>
<%Page.Execute%>
	<BODY id="CLASPBody">
		<h3>Classic ASP Framework - Global Administration Page</h3>
		<hr>
		<table ID="Table1">
			<tr><td><B>CLASP Started At:</B></td><td><%=Application("CLASP_Start")%></td></tr>
			<tr><td><B>Hit Count:</B> </td><td><%=Application("SC")%></td></tr>
			<tr><td><B>Page Hits:</B></td><td><%=Application("CLASP_PH")%></td></tr>
			<tr><td><B>Last Hit:</B></td><td><%=Application("CLASP_LAST_HIT")%></td></tr>
		</table>
		<%Page.OpenForm%>
			<table>
				<tr><td>Enable Trace:</td><td><%chkEnableTrace%></td></tr>
				<tr><td>Enable Debug:</td><td><%chkEnableDebug%></td></tr>
				<tr><td>Last Refresh:</td><td><%lblLastRefresh%></td></tr>
				<tr><td><a href="javascript:clasp.form.doPostBack('refresh','Page')">Refresh</a> | 
					    <a href="#" onclick="javascript:window.open('http://clasp.csharpjunkie.com/buttons.asp')">Open WebSite</a></td><td></td></tr>
			</table>
		<%Page.CloseForm%>
	</BODY>
	<%tmrTimer%>
</HTML>
<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim chkEnableTrace
	Dim chkEnableDebug
	Dim lblLastRefresh
	Dim tmrTimer
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		Set chkEnableTrace = New_ServerCheckBox("chkEnableTrace")
		Set chkEnableDebug = New_ServerCheckBox("chkEnableDebug")
		Set lblLastRefresh = New_ServerLabel("lblLastRefresh")
		Set tmrTimer	   = New_ServerTimer("tmrTimer",60000) 'defaults to 1000 millseconds		
	End Function

	Function Page_Controls_Init()
		If Application("SC") = "" Then
			Application("SC") = 1
		End If
		
		chkEnableTrace.AutoPostBack = True
		chkEnableDebug.AutoPostBack = True
		
		If Application("CLASP_DEBUG_ENABLED") <> "" Then
			Application("CLASP_DEBUG_ENABLED") = False
		End If

		If Application("CLASP_TRACE_ENABLED") <> "" Then
			Application("CLASP_TRACE_ENABLED") = False
		End If				
		
		chkEnableTrace.Checked = IIF(Application("CLASP_TRACE_ENABLED")<>"",Application("CLASP_TRACE_ENABLED"),False)
		chkEnableDebug.Checked = IIF(Application("CLASP_DEBUG_ENABLED")<>"",Application("CLASP_DEBUG_ENABLED"),False)
		Page.ShowRenderTime = chkEnableTrace.Checked
		Page.DebugEnabled = chkEnableDebug.Checked	
	End Function

	Function Page_Load()
		lblLastRefresh.Text =  Now
	End Function

'WEBCONTROLS POSTBACK EVENT HANDLERS	
	Function chkEnableTrace_Click()
		Application("CLASP_TRACE_ENABLED") = chkEnableTrace.Checked
		Page.ShowRenderTime = chkEnableTrace.Checked
	End Function


	Function chkEnableDebug_Click()
		Application("CLASP_DEBUG_ENABLED") = chkEnableDebug.Checked	
		Page.DebugEnabled = chkEnableDebug.Checked	
	End Function
	
	Function tmrTimer_OnTimer()
		chkEnableTrace.Checked = IIF(Application("CLASP_TRACE_ENABLED")<>"",Application("CLASP_TRACE_ENABLED"),False)
		chkEnableDebug.Checked = IIF(Application("CLASP_DEBUG_ENABLED")<>"",Application("CLASP_DEBUG_ENABLED"),False)
	End Function

'SUPPORTING FUNCTIONS (IF ANY)
%>
