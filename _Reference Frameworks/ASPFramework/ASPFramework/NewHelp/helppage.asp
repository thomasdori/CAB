<!--#Include File = "HelpController.asp" -->
<!--#Include File = "..\Controls\Server_TreeView.asp" -->
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<SCRIPT language="javascript">
	function ShowPage(url) {
		var obj = clasp.getObject("PageContents")
		obj.src = url;
		return false;
	}
</SCRIPT>
<%
	mLeftSectionTitle = "CLASP 2.0"
	Call PageController.RenderPage()	

':: PAGE CONTROLLER EVENT
Public Function Page_RenderForm()
%>
<TR height="<%=Browser.ScreenHeight - 290%>">
	<TD style="BORDER-RIGHT: #3366cc 1px solid" width="200" valign="top">
	<div id="HelpSection" style="OVERFLOW: auto; POSITION: relative; HEIGHT: <%=Browser.ScreenHeight - 290%>">
    <%tvwHelp%>
    </div>
    </TD>
   <TD width="*" valign="top">
	   <IFRAME style="border:dotted 1px #3366cc"  frameborder =0 Id = "PageContents" src="<%=mFrameURL%>" width='100%' height='100%'></IFRAME>
  </TD>
</TR>
<%End Function

':: PAGE VARIABLES
	Dim mFrameURL

':: PAGE CONTROLS
	Dim tvwHelp

':: PAGE EVENTS
	Public Function Page_Init
		mFrameURL  = "..\help\OverView.htm"
'		Page.DebugEnabled=True
		If "" & Application("TVHELP") = "" Then 
			'LoadFileToCache Server.MapPath("\ASPFramework\NewHelp\help.xml"),"TVHELP"		Dim oXML
			Dim oXML
			Set oXML = CreateObject(CLASP_DOM_PARSER_CLASS)			
			oXML.Load(Server.MapPath("\ASPFramework\NewHelp\help.xml"))
			Application("TVHELP") = oXML.Xml
			set oXML = Nothing			 			
		End If
		Set tvwHelp = New_ServerTreeView("tvwHelp","TVHELP")
		tvwHelp.MenuStyle = "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "treeview/help.js'></script>"
		tvwHelp.MenuStyleName = "tvTree_Help"
	End Function

	Public Function Page_LoadViewState()
		mFrameURL = Page.ViewState.GetValue("FrameURL")
	End Function

	Public Function Page_SaveViewState()
		Page.ViewState.Add "FrameURL",mFrameURL
	End Function

':: PAGE POSTBACK HANDLERS
	Public Function tvwHelp_OnClick(node) 
		If node.GetAttribute("F")<> "" Then
			mFrameURL = node.GetAttribute("F")
		End If
	End Function

':: PAGE HELPER ROUTINES
Function LoadFileToCache(sFileName,sCacheName)
	Dim oFSO
	Dim oTS
	Dim sContents
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oTS = oFSO.OpenTextFile(sFileName,1)
	sContents = oTS.ReadAll		
	oTS.Close	
	Set oTS  = Nothing
	Set oFSO = Nothing  	
	Application(sCacheName) = sContents
End Function

Public Function Page_LogOut(e)
	Response.Cookies("CLASP_SiteUser")  = ""
	Response.Cookies("CLASP_SiteEmail") = ""
End Function

Public Function Page_OnAccCreated(e)
	Select Case e.Instance
		Case "1"
			mFrameURL = "..\help\OverView.htm"	
		Case Else 
			mFrameURL = "Downloads.asp"	
	End Select
End Function
%>