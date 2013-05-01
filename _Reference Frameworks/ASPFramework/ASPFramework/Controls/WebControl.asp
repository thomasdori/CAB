<%Option Explicit%>
<!--#Include File = "StringBuilder.asp"-->
<!--#Include File = "Server_BrowserAgent.asp"-->
<!--#Include File = "CLASP_Setup.asp"-->
<%
'*********************************************************************************************************************
CONST VIEW_STATE_MODE_CLIENT = 1
CONST VIEW_STATE_MODE_SERVER = 2

Dim Browser
Set Browser = Nothing
Set Browser = New BrowserAgent

Dim Page
Set Page = Nothing 'Yeps, yeah have to do it before!
Set Page = New cPage

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::MAIN PAGE CONTROLER
	Public Sub Main()
		If Page.IsAJAXEvent Then
			Page.HandleServerEvent "Page_Authenticate_Request"
			Page.HandleServerEvent "Page_Authorize_Request"								
			Page.HandleServerEvent "Page_Init"
			Page.HandleServerEvent "Page_LoadViewState"
			Page.HandleServerEvent "Page_Load"
			Page.HandleServerEvent "Page_ProcessAjaxForm"
			Page.HandleServerEvent "Page_ClientEvent"
			Exit Sub		
		End If
	
		'load clasp core javascript files
		Response.Write vbNewLine & "<script language='javascript'>var SCRIPT_LIBRARY_PATH = """ & SCRIPT_LIBRARY_PATH & """;</script>" 
		Response.Write vbNewLine & "<script language='javascript' src='" & SCRIPT_LIBRARY_PATH & "claspx.js'></script>"

		Page.RenderClienStartuptScripts	
		Page.HandleServerEvent "Page_Authenticate_Request"
		Page.HandleServerEvent "Page_Authorize_Request"								
		Page.HandleServerEvent "Page_Init"

		If Not Page.IsPostBack Then
			Page.HandleServerEvent "Page_Controls_Init"
		End If
		
		If Page.IsPostBack Then
			Page.HandleServerEvent "Page_LoadViewState"
			Page.HandleServerEvent "ProcessPostBackData"
		End If
		
		Page.HandleServerEvent "Page_Load"

		If Page.IsPostBack Then
			Page.HandleServerEvent "RaiseChangedEvents"
		End If
		
		If Page.IsPostBack Then			
			Page.TraceCall Page.Control,"Handle PostBack -Start"
			Page.HandlePostBack ""
			Page.TraceCall Page.Control,"Handle PostBack -End"			
		End If		
		
		Page.HandleServerEvent "Page_PreRender"
		Page.HandleServerEvent "Page_SaveViewState"

		If Page.IsTemplatePresent Then
			Page.ApplySkin
		End If

	End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::EVENT
   Class ClientEvent
		Dim EventName
		Dim Source
		Dim Instance
		Dim CommandSource
		Dim ExtraMessage
		Dim TargetObject
		Public Property Get EventFnc()
			EventFnc = Source & "_" & EventName
		End Property
	End Class
	
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::PROPERTY BAG WRAPPER
	Class PropertyBag
		Dim Node
				
		Private Sub Class_Initialize()
			Set Node = Nothing
		End Sub
		
		Public Function Write(Name,Value)
			WriteProperty Name,Value,Node
		End Function

		Public Function WriteEx(Name,Value,Current)
			If Value <> Current Then
				WriteProperty Name,Value,Node
			End If
		End Function

		Public Function Item(x)
			Item = ReadPropertyByIndex(x,Node)
		End Function

		Public Function ReadEx(Name,Def)		
			ReadEx = ReadProperty(Name,Node)
			If ReadEx = "" Then
				ReadEx  = Def
			End If			
		End Function
		
		Public Function Read(Name)		
			Read = ReadProperty(Name,Node)
		End Function

		Public Function ReadBoolean(Name)		
			ReadBoolean = CBool(ReadProperty(Name,Node) & "0")
		End Function

		Public Function ReadInt(Name)		
			ReadInt = ReadProperty(Name,Node)
			If ReadInt = "" Then
				ReadInt = 0
			Else
				ReadInt = CInt(ReadInt)
			End If
		End Function

		Public Function ReadDouble(Name)		
			ReadDouble = ReadProperty(Name,Node)
			If ReadDouble = "" Then
				ReadDouble = 0
			Else
				ReadDouble = CInt(ReadDouble)
			End If			
		End Function
		
		Public Property Get Count
			Count = node.attributes.length			
		End Property		

	End Class
	
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::CONTROLS COLLECTION
   Class ControlsCollection
		Private mintIndex
		Private mvarControlArray
		Public  GrowRate		

		Private Sub Class_Initialize()
			mintIndex = 0
			GrowRate = 5
		End Sub
		
		Public Property Get Count
			Count = mintIndex
		End Property
		
		Public Sub Clear()
			mintIndex = 0
		End Sub
		
		Private Sub ReSize()
			Dim x,mx			
			If mintIndex  = 0 Then			
				Redim mvarControlArray(GrowRate)
			Else
				Redim Preserve mvarControlArray(Ubound(mvarControlArray) + GrowRate)
			End If
			mx = Ubound(mvarControlArray)
			For x = mintIndex To mx
				Set mvarControlArray(x) = Nothing
			Next
		End Sub
		
		Public Function Item(NameOrIndex)		  
			If IsNumeric(NameOrIndex) Then
				Set Item = mvarControlArray(CInt(NameOrIndex)) 
			Else
				Set Item = GetByName(NameOrIndex)
			End If
		End Function

		Public Function GetByName(name)
			Dim X,MX,TheControl
			mx = mintIndex - 1	
			Set GetByName = Nothing
			For x = 0 To mx				
				Set TheControl = mvarControlArray(x)
				If  TheControl.ControlID = name Then
					Set GetByName = TheControl
					Exit For
				End If
			Next
		End Function
	
		Public Sub Add(obj)
			If mintIndex = 0 Then				
				Call ReSize()
			Else
				If mintIndex > UBound(mvarControlArray) Then
				   Call Resize()
				End If
			End If		
			Set mvarControlArray(mintIndex) = obj
			mintIndex = mintIndex + 1
		End Sub
		
		Public Sub Remove(objRemove)
			Dim x,mx,obj
			mx = mintIndex - 1
			For x = 0 To mx
				Set obj = mvarControlArray(x)
				If Not obj Is Nothing Then
					If objRemove Is obj Then
						Set mvarControlArray(x) = Nothing
						If x = mintIndex Then						
							 mintIndex = mintIndex - 1 
						Else
							mintIndex = mintIndex - 1
							Set  mvarControlArray(x) = 	  mvarControlArray(mintIndex)
							Set mvarControlArray(mintIndex) = Nothing
						End If
						Exit Sub
					End If	
				Else
					Exit Sub
				End If
			Next			
		End Sub 		
	End Class
	
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::WEB CONTROL
   Class WebControl	
		Dim Parent
		Dim Owner
		Dim EnableViewState 
		Dim ControlID
		Dim Enabled
		Dim Visible    
		
		Dim Style		
		Dim Attributes
		Dim CssClass
		Dim TabIndex
		Dim ToolTip
		Dim DataTextField
		Dim IsViewStateRestored
		Dim ImplementsOnInit
		Dim ImplementsOnLoad
		Dim ImplementsProcessPostBack
		Dim ViewState
		Dim OverrideTemplate
		Dim SupportsClientSideEvent		
		Private Sub Class_Initialize()
			Set Owner   = Page 'ARU U SURE??? Nothing
			Set Parent  = Nothing			
			ControlID   = ""
			Style       = ""
			Enabled     = True
			Visible     = True
			TabIndex    = 0
			DataTextField = ""
			IsViewStateRestored = False
			Set ViewState       = Nothing
			ImplementsOnInit    = False
			ImplementsOnLoad    = False
			ImplementsProcessPostBack = False
			OverrideTemplate	= False
			SupportsClientSideEvent = False
			
			If Not Page Is Nothing Then
				EnableViewState = Page.Control.EnableViewState
				Page.Controls.Add Me
				Set Parent = Page.Control
			Else
				EnableViewState  = True 'ViewState is enabled by default.
			End If
		End Sub
		
		Private Sub Class_Terminate()
			Set ViewState = Nothing
		End Sub
				
		Public Property Get IsVisible()   
			If Visible Then
				If Parent Is Nothing Then 
					IsVisible = True
				Else
					IsVisible = Parent.IsVisible
				End If		
			Else
				IsVisible = False
			End If
		End Property

		Public Property Let Name(ByVal v)
			ControlID = v
			If Page.IsPostBack Then
				If Page.IsViewStateLoaded Then
					If Not IsViewStateRestored Then
						Call RestoreViewState()
					End If
				End If
			End If
		End Property
		
		Public Property Get Name()
			Name = ControlID
		End Property
		
		Private Function RestoreViewState()
			Dim xml,vs,PageNode,IsNew
			Page.ViewState.GetDomObject xml
			Set PageNode  = GetSection("PAGE",       xml.ChildNodes(0))
			Set vs		  = GetSectionEx(ControlID,  PageNode.Childnodes(0) , IsNew)
			If IsNew Then				
				WriteProperties vs
				PageNode.ChildNodes(0).AppendChild(vs)
			End If

			If Not IsNew Then
				ReadProperties vs
			End If

			If ImplementsProcessPostBack Then				
				Call Owner.ProcessPostBack()
			End If											
						
			If Page.IsViewStateSaved Then				
				WriteProperties vs
			End If
			
			Set xml = Nothing
			Set vs  = Nothing
			Set PageNode = Nothing
		End Function

		Public Function ReadProperties(vs)
			Dim wc
			Dim bag
			Dim IsNew
			Dim varData
			IsViewStateRestored = True			

			Set wc = GetSectionEx("WC",vs,IsNew)			
			
			If Not IsNew Then
				CssClass   = ReadProperty("C",wc)
				Style      = ReadProperty("S",wc)
				Attributes = ReadProperty("A",wc)
				
				varData    = Split(ReadProperty("P",wc),";")
				
				Visible = CBool(varData(0))
				Enabled = CBool(varData(1))
				EnableViewState = CBool(varData(2))
				TabIndex = CInt("0" & varData(3))
				OverrideTemplate = CBool(varData(4))
				SupportsClientSideEvent = CBool(varData(5))
									
				Set bag = New PropertyBag
				Set ViewState = bag
				Set bag.Node = vs
				Owner.ReadProperties(bag)
				Set bag = Nothing
			End If
						
		End Function
		
		Public Function WriteProperties(vs)
			Dim varState
			Dim wc
			Dim bag
						
			If vs Is Nothing Then
				Exit Function
			End If

			Set wc = GetSection("WC",vs)
			WriteProperty "C", CssClass,wc
			WriteProperty "S", Style,wc
			WriteProperty "A", Attributes,wc			
			WriteProperty "P", Cint(Visible) & ";" & CInt(Enabled) & ";" & CInt(EnableViewState) & ";" & TabIndex & ";" & CInt(OverrideTemplate) & ";" & CInt(SupportsClientSideEvent),wc
			
			If 	ViewState Is Nothing Then
				Set bag = New PropertyBag
				Set bag.Node = vs
				Set ViewState = bag
			Else
				Set bag = ViewState
			End If
			
			Owner.WriteProperties(bag)
			vs.AppendChild wc
			Set bag = Nothing
				
		End Function

		Public Sub HandleServerEvent(evt)
			If Not Owner Is Nothing Then
				Owner.HandleServerEvent(evt)
			End If
		End Sub
		
		Public Function HandleClientEvent(evt)
			If Parent Is Page.Control Then
				HandleClientEvent = Owner.HandleClientEvent(evt)
			Else
				If evt.TargetObject.Parent Is Me Then
					HandleClientEvent = Owner.HandleClientEvent(evt)
				Else
					Page.TraceImportantCall Me, "Bubble Event To " & Parent.Name
					HandleClientEvent = Parent.HandleClientEvent(evt)
				End If
			End If
		End Function
		
	End Class

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::PAGE CONTROL
	 Class cPage
	
		Dim Controls
		Dim Control
		Dim PageID
		Dim PreserveViewState	
		Private mobjRenderTrace
		Private mdteStartRender
		Private mdteEndRender
		Private mvarRenderTrace
		Private mvarRenderTraceRow
		Private mobjStackTrace
		Private mvarStackTraceRow
		Private mobjClientImages
		Private mobjClientScripts
		Private mobjClientStartup		
		Private mobjErrorTrace
		Private mPostBackEvent

		Private mStartPageTimer				
		Private mintTabIndex
		Private mvarPostBackHandlers
		Private mintPostBackHandlerIndex
		Private mLoadedLibraries
		
		Dim CompressFactor
		Dim IsCompressed				
		Dim IsRedirectedPostBack		
		Dim IsPostBack
		
		Dim Action
		Dim ActionSource
		Dim ActionSourceInstance
		Dim ActionXMessage				
		Dim DebugEnabled		
		Dim DebugLevel
		Dim ShowRenderTime
		Dim ViewStateMode
		Dim ViewState
		Dim PageFormAction
		Dim AutoResetFocus
		Dim AutoResetScrollPosition
		Dim OnSubmitStatement
		Dim ScrollTop
		Dim ScrollLeft
		Dim IsViewStateLoaded
		Dim IsViewStateSaved		
		Dim IsTemplatePresent
		Dim IsAJAXEvent
		Dim AjaxXml
		Private Sub Class_Initialize()
			
			mdteStartRender = Now
			mStartPageTimer = Timer()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 
			Set Controls = New ControlsCollection
			    Controls.GrowRate = 50
			
			Set mobjStackTrace    = New StringBuilder
			Set mobjRenderTrace   = New StringBuilder
			Set mobjErrorTrace    = New StringBuilder
			Set mLoadedLibraries  = New StringBuilder
			Set mobjClientScripts = Nothing
			Set mobjClientStartup = Nothing			
			Set mobjClientImages  = Nothing
			Set mPostBackEvent    = Nothing			
		
			IsAJAXEvent       = False	
			IsTemplatePresent = False			
			IsViewStateLoaded = False
			IsViewStateSaved  = False
			AutoResetScrollPosition = True
			AutoResetFocus = False
			DebugEnabled   = False
			DebugLevel	   = 3' 3) Display everything. 2) Display Controls Tree, Stack Trace, Forms Collection, etc. 1) For Postback data only
			Control.EnableViewState = True 
			mintTabIndex = 1
			CompressFactor = -1																		

			PageID = ""	
			PreserveViewState = False
			
			If Application("CLASP_DEBUG_ENABLED")<> "" Then
				DebugEnabled = Application("CLASP_DEBUG_ENABLED")
			End If
			
			If Application("CLASP_TRACE_ENABLED")<> "" Then
				ShowRenderTime = Application("CLASP_TRACE_ENABLED")
			End If
			ShowRenderTime=False
						
			mvarRenderTraceRow = 0		
			mvarStackTraceRow  = 0
			Control.ControlID = "Page"						
			
'			ViewStateMode = VIEW_STATE_MODE_CLIENT
			ViewStateMode = VIEW_STATE_MODE_SERVER
			
			Set ViewState = New_ViewStateObject()
			
			If Request.Form("AjaxPost") = "1" Then				
				IsAJAXEvent  = True
				Action       = Request.Form("AjaxAction")
				ActionSource = Request.Form("AjaxSource")
				
				If Request.Form("AjaxMode") = "1" Then
					Response.Clear
					ExecuteEventFunctionParams Action,ActionSource
					Response.End
				End If
				
				Set AjaxXml  = CreateObject(CLASP_DOM_PARSER_CLASS)			
				If Not AjaxXml.LoadXml(Request.Form("AjaxForm")) Then
					Response.Clear
					Response.Write "Error Processing the AJAX Form" & Server.HTMLEncode(Request.Form("AjaxForm"))
					Response.End
				End If
				TraceMsg "Initialize Page Object","Render Time: " & FormatNumber( Timer()-mStartPageTimer,6)
				Exit Sub
			End If

			IsRedirectedPostBack = (Request.Form("__ISREDIRECTEDPOSTBACK") <> "")						
			IsPostBack           = (Request.Form("__ISPOSTBACK")= "True") And Not IsRedirectedPostBack
			
			If IsPostBack Or IsRedirectedPostBack Then
				Action		         = Request.Form("__ACTION")
				ActionSource         = Request.Form("__SOURCE")
				ActionSourceInstance = Request.Form("__INSTANCE")
				ActionXMessage       = Request.Form("__EXTRAMSG")
			End If				
			
			If IsPostBack Then					
				IsCompressed = CBool(IIF(Request.Form("__VIEWSTATECOMPRESSED")="","False",Request.Form("__VIEWSTATECOMPRESSED")))
				ScrollTop    = Request.Form("__SCROLLTOP")				
				ScrollLeft	 = Request.Form("__SCROLLLEFT")				
			Else
				Session("CLASP_VS") = ""
				IsPostBack   = False
				IsCompressed = False
				ScrollTop    = 0
				ScrollLeft	 = 0
				Call Page_WriteViewState()				
			End If
			TraceMsg "Initialize Page Object","Render Time: " & FormatNumber( Timer()-mStartPageTimer,6)
		End Sub
		
		Private Sub Class_Terminate()
			mdteEndRender = Now
			HandleServerEvent("Page_Terminate")			
			If ShowRenderTime Then
				Response.Write "<span style='color:navy;font-size:10pt;font-family:tahoma;font-weight;width:100%;background-color:#DDDDDD'><b>Total Render Time:</b> "  & FormatNumber( Timer()-mStartPageTimer,6) & "</span><br>"
			End If
			If DebugEnabled Then				
				If Not ShowRenderTime Then
					Response.Write "<span style='color:navy;font-size:10pt;font-family:tahoma;font-weight;width:100%;background-color:#DDDDDD'><b>Total Render Time:</b> "  & FormatNumber( Timer()-mStartPageTimer,6) & "</span><br>"
				End If
				If IsPostBack Or IsRedirectedPostBack Then
					OpenDebugSection "PostBack Event Datta"
						WriteToDebugSection "IsPostBack",IsPostBack
						WriteToDebugSection "Is RedirectedPostBack",IsRedirectedPostBack
						WriteToDebugSection "Event",Request.Form("__ACTION")
						WriteToDebugSection "Object",Request.Form("__SOURCE")
						WriteToDebugSection "Instance",Request.Form("__INSTANCE")
						WriteToDebugSection "Extra Message",Request.Form("__EXTRAMSG")
					CloseDebugSection
				End If				
					
				PrintIISCollectionFor Request.Form,"Form Collection"
				PrintIISCollectionFor Request.Cookies,"Cookies Collection"
				PrintIISCollectionFor Request.QueryString,"Query String"
				PrintIISCollectionFor Application.Contents,"Application Contents"				
				PrintIISCollectionFor Session.Contents,"Session Contents"

				OpenDebugSection "Session Information"
					WriteToDebugSection "ID",Session.SessionID
					WriteToDebugSection "Timeout",Session.Timeout
					WriteToDebugSection "CodePage",Session.CodePage
				CloseDebugSection

				OpenDebugSection "Response Information"
					WriteToDebugSection "Content Type",Response.ContentType
					WriteToDebugSection "Cache Control",Response.CacheControl
					WriteToDebugSection "Status",Response.Status
				CloseDebugSection
					
				Response.Write("<hr><br><span style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro;width:100%'>Errors Trace</span><br><table width='100%' cellpadding=0 cellspacing=0 style='font-family:tahoma'><tr style='background-color:black;color:white;font-family:tahoma'><td>Control</td><td>Msg</td><td>Severity</td></tr>")
					Response.Write(mobjErrorTrace.ToString())
				Response.Write("</table><hr>")		
				
				If Page.DebugLevel > 1 Then

					'Prints the stack trace
					Response.Write("<br><Span style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro;width:100%'>Calls Trace</span><br><table width='100%' cellpadding=0 cellspacing=0 style='font-family:tahoma;font-size:8pt'><tr style='background-color:black;color:white;font-family:tahoma'><td>Control</td><td>Msg</td><td>Time</td></tr>")
					Response.Write(mobjStackTrace)
					Response.Write("</table><hr>")

					'Prints the Render Trace
					Response.Write("<br><Span style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro;width:100%'>Trace Output</span><br><table width='100%' cellpadding=0 cellspacing=0 style='font-family:tahoma;font-size:8pt'><tr style='background-color:black;color:white;font-family:tahoma'><td>Control</td><td>From</td><td>To</td></td><td>Total</td></tr>")
					Response.Write( "<tr style='font-weight:bold'><td>Page</td><td>" & mdteStartRender & "</td><td>" &  mdteEndRender & "</td><td>" & DateDiff("s",mdteStartRender,mdteEndRender )  & "</td></tr>" )
					Response.Write(mobjRenderTrace)
					Response.Write("</table><hr>")		
					
					'Prints the Loaded Libraries
					Response.Write "<br><span style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro;width:100%'>Loaded Libraries</span><br><span style='font-size:8pt'>" & mLoadedLibraries.ToString2("<br>") & "</span><hr>"
					
					'Prints the Controls Tree
					Call GetControlsMap()
				End If

				If Page.DebugLevel > 2 Then
					PrintIISCollectionFor Request.ServerVariables,"Server Variables"
					
					OpenDebugSection "View State Information"
						WriteToDebugSection "Size",Len(Page.ViewState.GetViewState)
						WriteToDebugSection "Is Compressed",Page.IsCompressed
						WriteToDebugSection "Compress Factor",Page.CompressFactor & " (in bytes)"
						If Page.IsCompressed Then
							WriteToDebugSection "Compressed Size",Len(Page.ViewState.GetViewStateBase64(Page.CompressFactor))
						End If
					CloseDebugSection									
					Response.Write "<br><textarea rows=20  style='width:100%' >" & Page.ViewState.GetViewState & "</textarea>"
				End If
								
			End If			

			Set ViewState = Nothing					
		End Sub

		Private Sub OpenDebugSection(SectionName)
				Response.Write "<br><table width=100% border=0 cellspacing=0>"
				Response.Write "<tr style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro'><td  colspan=2>" & SectionName & "</td></tr>"
				Response.Write "<tr style='font-size:8pt;font-family:tahoma;background-color:Black;color:white'><td>Property</td><td>Value</td></tr>"
		End Sub

		Private Sub CloseDebugSection()
				Response.Write "</table>"
		End Sub

		Private Sub WriteToDebugSection(propName,propValue)
				Response.Write "<tr><td width=25% nowrap style='color:navy;font-size:8pt;font-weight:bold'>"
				Response.Write propName
				Response.Write "</td><td style='font-size:8pt;'>" & propValue & "</td></tr>"
		End Sub

		Private Function PrintIISCollectionFor(obj,title)
				Dim x,varKey,tName
				If obj.Count = 0  Then
					Exit Function
				End If
				Response.Write "<br><Span style='font-size:10pt;font-family:tahoma;font-weight:bold;background-color:Gainsboro;width:100%'>" & title & "</span><br>"
				Response.Write "<table width='100%' cellpadding=2 cellspacing=0 style='font-family:tahoma'><tr style='background-color:black;color:white;font-family:tahoma;font-size:10pt'><td>Item</td><td>Value</td></tr>"
				For x = 1 To obj.Count
					varKey = obj.Key(x)
					If varKey<> "__VIEWSTATE"  Then
						Response.Write "<tr><td valign=top style='font-size:8pt;color:navy;font-weight:bold'width=25% nowrap  >"  &  varKey & "</td>"						
						If IsObject(obj.Item(varKey) ) Then
							tName = TypeName(obj.Item(varKey) ) 
							If tName = "IStringList" Then
								tName = obj.Item(varKey)
							End If
							Response.Write "<td valign=top style='font-size:8pt'>"  & tName  & "</td></tr>"
						Else
							Response.Write "<td valign=top style='font-size:8pt'>"  & Server.HTMLEncode( obj.Item(varKey) ) & "</td></tr>"
						End If
					End If
				Next
				Response.Write("</TABLE><hr>")		
		End Function

		Public Function GetViewStateID()	
			'the view stateid is a combination of PageID and SessionID. Unique viewstates can be stored in the session object by just setting the PageID property
			GetViewStateID = PageID & Session.SessionID
		End Function
		
		Public Function RegisterLibrary(LibName)		
			If mLoadedLibraries.FindString(LibName) < 0 Then
				mLoadedLibraries.Append LibName
			End If
		End Function
		
		Public Function IsLibraryLoaded(LibName)
			IsLibraryLoaded = ( mLoadedLibraries.FindString(LibName) >= 0 )
		End Function

		Public Function CheckLibDependency(LibName,RequiredBy)		
			If Not Page.IsLibraryLoaded(LibName) Then
				Response.Write "<b>" & RequiredBy & "</b>: Library <font color=blue>" & LibName & "</font> is missing. Add it before this library.<br>"
				Response.End
			End If
		End Function

		Public Function ApplySkin
			Dim x,mx,ctrl,skin			
			Dim arrFnc(20,1)
			mx = Controls.Count-1
			For x = -1 To mx
				If x<0 Then
					Set ctrl = Page.Control
				Else
					Set ctrl = Controls.Item(x)
				End If
				If ctrl.OverrideTemplate = False Then
					Set skin = GetControlSkin(ctrl,arrFnc)
					If Not skin Is Nothing Then
						Call skin(ctrl)
					End If
				End If
			Next
		End Function	
		
		Private Function GetControlSkin(ctrl,arrFnc) 
			Dim LibName
			Dim x,mx
			
			LibName = TypeName(ctrl.Owner)
			mx = UBound(arrFnc)	
			For x = 0 To mx
				If arrFnc(x,0) = LibName Then
					Set GetControlSkin = arrFnc(x,1)
					Exit Function
				ElseIf arrFnc(x,0) = "" Then
					Set arrFnc(x,1) = GetFunctionReference(LibName & + "_Skin")
					arrFnc(x,0) = LibName
					Set GetControlSkin = arrFnc(x,1)
					Exit Function
				End If							
			Next
			Set GetControlSkin = Nothing					
		End Function


		Public Function GetControlsMap()
			Dim x,mx
			Dim objControl
			OpenDebugSection "Page Controls"
			mx = Page.Controls.Count - 1
			WriteToDebugSection "Page",TypeName(Page)
			For x = 0 To mx 
				Set objControl   = Page.Controls.Item(x)				
				'If Not objControl.Parent Is Nothing Then
					WriteToDebugSection "--" & objControl.Name ,TypeName(objControl.Owner)' & " , Parent:" & objControl.Parent.Name
				'Else
					'WriteToDebugSection objControl.Name,"--"
				'End If
			Next
			CloseDebugSection									
		End Function

		
		Private Function LoadViewState()
			Dim vsFnc
			Set vsFnc = GetFunctionReference("Page_LoadPageStateFromPersistenceMedium")
			If vsFnc Is Nothing Then
				Page.TraceCall Page.Control, "Load From Persistence Medium"
				If ViewStateMode = VIEW_STATE_MODE_SERVER Then 'Session("CLASP_VS") <> "" Then
					Call ViewState.LoadViewState(Session(GetViewStateID()))
				Else
					'If from Ajax then do load from ajax request
					If IsAJAXEvent Then
						Call ViewState.LoadViewStateBase64(GetAjaxViewState(),IsCompressed)
					Else
						Call ViewState.LoadViewStateBase64(Request.Form("__VIEWSTATE"),IsCompressed)
					End If
					
				End If
			Else
			   Page.TraceCall Page.Control, "Load From Persistence Medium - Custom"
			   Call vsFnc()
			End If						
		End Function
		
		Private Function SaveViewState()
		    Dim vsFnc
		    Dim bolIsCompressed			
			bolIsCompressed = FALSE			
			Set vsFnc = GetFunctionReference("Page_SavePageStateToPersistenceMedium")
			If vsFnc Is Nothing Then
			   Page.TraceCall Page.Control, "Save From Persistence Medium"
			   If Page.ViewStateMode = VIEW_STATE_MODE_CLIENT  Then
					If IsAjaxEvent Then
						Response.Write "|*|" & ViewState.GetViewStateBase64(CompressFactor,bolIsCompressed)
					Else
						Response.Write "<input type='hidden' name='__VIEWSTATE' id='__VIEWSTATE' value='"
						Response.Write ViewState.GetViewStateBase64(CompressFactor,bolIsCompressed)
						Response.Write "'>"
						Response.Write vbNewLine
					End If
				Else
					Session(GetViewStateID()) = Page.ViewState.GetViewState()
				End If
				IsCompressed = bolIsCompressed
				Page.TraceImportantCall Page.Control, "View State Compressed: " & Page.IsCompressed			
			    Response.Write vbNewLine
			Else
			   Page.TraceCall Page.Control, "Save From Persistence Medium - Custom"
			   Call vsFnc()
			End If	
		End Function

		Public Sub RegisterClientImage(ImageName,ImageUrl)			
			If mobjClientImages Is Nothing Then
				 Set mobjClientImages = New_ViewStateObject()
			End If			
			mobjClientImages.Add ImageName,ImageUrl			
		End Sub
					
		Public Sub RegisterClientScript(ScriptName,ScriptText)						
			If mobjClientScripts Is Nothing Then
				 Set mobjClientScripts = New_ViewStateObject()
			End If			
			mobjClientScripts.Add ScriptName,ScriptText			
		End Sub

		Public Sub RegisterClientStartupScript(ScriptName,ScriptText)			
			If mobjClientStartup Is Nothing Then
				 Set mobjClientStartup = New_ViewStateObject()
			End If			
			mobjClientStartup.Add ScriptName,ScriptText			
		End Sub
				
		Public Function RegisterEventListener(object,eventName,fncName) 
			Dim objName
			Dim script			
			Select Case TypeName(object) 
				Case "WebControl"
					objName = object.ControlID
				Case "String"
					objName  = object
				Case Else
					objName = object.Control.ControlID
			End Select			
			script = vbNewLine &  "clasp.events.addListener('" & objName & "','" & eventName & "','" & replace(fncName,"'","\'") & "');"			
			script  = "<script language='JavaScript'>" & script &  vbNewLine & "</script>"
			If mobjClientScripts Is Nothing Then
				Page.RegisterClientScript objName & "_0" ,script
			Else
				Page.RegisterClientScript objName & "_" & mobjClientScripts.Count ,script
			End If						
		End Function

		Private Function RenderClientImages()
			Dim x,mx
			Dim sStartTag
			If Not 	mobjClientImages Is Nothing Then
				Response.Write "<script language='JavaScript'>" + vbNewLine
				Response.Write "var cImg;"+ vbNewLine
				sStartTag = "cImg = new Image();cImg.src = '"
				mx  = mobjClientImages.Count - 1				
				For x = 0 To mx
					Response.Write sStartTag
					Response.Write mobjClientImages.GetValueByIndex(x)
					Response.Write "';" & vbNewLine
				Next				
				Set mobjClientImages = Nothing				
				Response.Write "</script>"
			End If			
		End Function
		
		Private Function RenderClientScripts()
			Dim x,mx			
			If Not 	mobjClientScripts Is Nothing Then
				mx  = mobjClientScripts.Count - 1				
				For x = 0 To mx
					Response.Write vbNewLine
					Response.Write mobjClientScripts.GetValueByIndex(x) 
				Next				
				Set mobjClientScripts = Nothing				
			End If			
		End Function
		
		Private Function Page_ProcessAjaxForm()
			Dim x,mx
			Dim ctrl
			Dim ctrlName
			Dim objControl			
			mx = AjaxXml.childNodes(0).childNodes.length- 1
			'response.Clear
			For x = 0 To mx				
				ctrlName = AjaxXml.childNodes(0).childNodes(x).nodeName
				Set ctrl = Page.Controls.Item(ctrlName)
				If Not ctrl is nothing then
					'Response.Write ctrl.ControlID
					If ctrl.SupportsClientSideEvent Then
						'On Error Resume Next						
						ctrl.Owner.SetValue UrlDecode(AjaxXml.childNodes(0).childNodes(x).text)
						If err.number>0 then
							Err.Clear
							On Error Goto 0
						End If
					End If
				End if
			Next		
			'response.End				
			Set objControl = Nothing
		End Function
		
		Private Function GetAjaxViewState()
			Dim objFields

			Set objFields = AjaxXml.selectNodes("//__VIEWSTATE")
			If objFields.length > 0	Then	
				GetAjaxViewState = replace(objFields.item(0).text," ","+") ' objFields.item(0).text
			Else
				Response.Clear
				Response.Write "ERROR: Invalid View State"
				Response.End
			End If
			Set objFields = Nothing
		End Function
				
		Public Function RenderClienStartuptScripts()
			Dim x,mx
			If Not 	mobjClientStartup Is Nothing Then
				mx  = mobjClientStartup.Count - 1				
				For x = 0 To mx
					Response.Write vbNewLine
					Response.Write mobjClientStartup.GetValueByIndex(x) 
				Next				
				Set mobjClientStartup = Nothing				
			End If			
		End Function
				
		Public Function GetNextTabIndex()
			GetNextTabIndex = mintTabIndex
			mintTabIndex = mintTabIndex + 1
		End Function

		Public Function GetAjaxEventScript(ctrl,callback,fullPost,JsEvent)
			if JsEvent = "href" then
				GetAjaxEventScript = " javascript:clasp.ajax.invoke(" & """" & ctrl.Control.ControlID & "_OnClientEvent" & """" & ",false," & callback & "," & """" &  ctrl.Control.ControlID & """"  & "," & iif(fullpost,"1","0") & ") "
			else
				GetAjaxEventScript = " onclick='clasp.ajax.invoke(" & """" & ctrl.Control.ControlID & "_OnClientEvent" & """" & ",false," & callback & "," & """" &  ctrl.Control.ControlID & """"  & "," & iif(fullpost,"1","0") & ")' "
			end if
		End Function

		
		Public Function GetEventScript(ClientTrigger, ObjectName, EventName, ByVal Instance, ExtraMessage)
			If Instance <> "this" Then
				Instance = "'" & Instance & "'"
			End If
			GetEventScript = " " & ClientTrigger & " = ""javascript:clasp.form.doPostBack('" & EventName  & "','" & ObjectName & "'," & Instance & ",'" & ExtraMessage  & "')"" "
		End Function
		
		Public Function GetEventScriptRedirect(ClientTrigger, ObjectName, EventName, ByVal Instance, ExtraMessage,Action)
			If Instance <> "this" Then
				Instance = "'" & Instance & "'"
			End If
			GetEventScriptRedirect = " " & ClientTrigger & " = ""javascript:clasp.form.doPostBack('" & EventName  & "','" & ObjectName & "'," & Instance & ",'" & ExtraMessage  & "','" & Action & "')"" "			
		End Function

		Public Function GetEventScriptNewWindow(ClientTrigger, ObjectName, EventName, ByVal Instance, ExtraMessage,Action,TargetWindow)
			If Instance <> "this" Then
				Instance = "'" & Instance & "'"
			End If
			GetEventScriptNewWindow = " " & ClientTrigger & " = ""javascript:clasp.form.doPostBack('" & EventName  & "','" & ObjectName & "'," & Instance & ",'" & ExtraMessage  & "','" & Action & "','" & TargetWindow & "')"" "
		End Function

		Public Function OpenWindow(Name,URL,Left,Top,Width,Height,Scrollable,Resizable,ExtraAttributes)
			Dim ScriptName
			ScriptName = "Window_" & Name
			RegisterClientScript ScriptName,"<script>" & GetOpenWindowScript(Name,URL,Left,Top,Width,Height,Scrollable,Resizable,ExtraAttributes) & "</script>"
		End Function
		
		Public Function GetOpenWindowScript(Name,URL,Left,Top,Width,Height,Scrollable,Resizable,ExtraAttributes)
			Dim ScriptText
			If Left = "" Then Left = 0
			If Top  = "" Then Top  = 0
			ScriptText = "clasp.openWindow('" &	 Name & "','" & Url & "'," & Left & "," & Top & "," & Width & "," & Height & "," & iif(Scrollable,"true","false") & "," & iif(Resizable,"true","false") & ",'" & ExtraAttributes & "')"	
			GetOpenWindowScript = ScriptText
		End Function
	
		
		Public Sub TraceRender(FromStart,ToEnd,varControl)
			Dim varValue
			
			If DebugEnabled Then
				If varControl = ActionSource Then
					varValue = "<span style='color:red;font-weight:bold'>(" & varControl & ")</span>"
				Else
					varValue = varControl
				End If
				mobjRenderTrace.Append  "<tr" & iif(mvarRenderTraceRow=0," style='background-color:lightblue'","") & "><td>"
				mobjRenderTrace.Append varValue & "</td><td>" & FromStart & "</td><td>" & ToEnd & "</td><td>" & DateDiff("s",FromStart,ToEnd ) & "</td></tr>" & vbNewLine
				mvarRenderTraceRow = 1 - mvarRenderTraceRow 
			 End If
		End Sub						
		
		Public Sub TraceCall(obj,msg)
			If DebugEnabled Then
				mobjStackTrace.Append "<tr " &  iif(mvarStackTraceRow=0," style='background-color:lightblue'","")  & "><td>"
				mobjStackTrace.Append obj.ControlID & "</td><td>" & msg & "</td><td>" & Now & "</td></tr>"
				mvarStackTraceRow = 1 - mvarStackTraceRow
			End If
		End Sub

		Public Sub TraceMsg(method,msg)
			If DebugEnabled Then
				mobjStackTrace.Append "<tr " &  iif(mvarStackTraceRow=0," style='background-color:lightblue'","")  & "><td style='font-weight:bold;coloir:blue'>"
				mobjStackTrace.Append method & "</td><td>" & msg & "</td><td>" & Now & "</td></tr>"
				mvarStackTraceRow = 1 - mvarStackTraceRow
			End If
		End Sub


		Public Sub TraceImportantCall(obj,msg)
			If DebugEnabled Then
				mobjStackTrace.Append "<tr " &  iif(mvarStackTraceRow=0," style='background-color:lightblue;color:red;font-weight:bold'"," style='color:red;font-weight:bold' ") & "><td>"	
				mobjStackTrace.Append obj.ControlID & "</td><td>" & msg & "</td><td>" & Now & "</td></tr>"
				mvarStackTraceRow = 1 - mvarStackTraceRow
			End If
		End Sub

		
		Public Sub TraceError(Control,Message,Severity)		
				If Err.number > 0 Then
					mobjErrorTrace.Append "<tr><td>" & IIF(Not Control Is Nothing,Control.ControlID,"") & "</td><td>" & err.Number & ":" & err.Description & "<br>" & Message & "</td><td>" & Severity & "</td></tr>" & vbNewLine
					Err.Clear
				End If
				Err.Clear
				On Error Goto 0 'Reset Error
		End Sub
		
		Public Function ReadProperties(bag)
			'Page.TraceCall Page.Control, "Page->Getting state" 
		End Function
		
		Public Function WriteProperties(bag)
			'Page.TraceCall Page.Control, "Page->Setting state" 
		End Function
		
		Private Function Page_WriteViewState()
			Dim x,mx
			Dim objChildNode	
			Dim objControl
			Dim obj
			Dim Node
			
			ViewState.GetDomObject obj
			Set Node = GetSection("PAGE",obj.ChildNodes(0)) '0			
			Control.WriteProperties(node)
			mx = Controls.Count	- 1
			For x = 0 To mx										
				Set objControl   = Controls.Item(x)
				Set objChildNode = GetSection(objControl.ControlID,node)					
				If objControl.EnableViewState Then
					objControl.WriteProperties(objChildNode)
				End If
				node.AppendChild objChildNode
			Next
			'If Not IsPostBack Then									
				Call obj.ChildNodes(0).AppendChild(Node) '0
			'End If
			Set objControl = Nothing
			Set objChildNode = Nothing
			Set obj = Nothing
		End Function

		Private Function Page_ReadViewState()
			Dim x,mx
			Dim objControl
			Dim objChildNode	
			Dim obj
			Dim node
			
			ViewState.GetDomObject obj						
			If obj.parseError.errorCode <> 0  Then
				With obj
					Response.Write "<br><b>ViewState failed to load:</b><hr>"
					Response.Write "Reason:" & .parseError.reason
					Response.Write "<br>errorCode:" & .parseError.errorCode
					Response.Write "<br>srcText:[" & .parseError.srcText & "]"
					Response.Write "<br>Line:" & .parseError.line
					Response.Write "<br>Col:" & .parseError.linepos
					Response.End
				End With			
			End If
						
			Set node = GetSection("PAGE",obj.ChildNodes(0)) '0
			mx = Page.Controls.Count - 1
			For x = 0 To mx 
				Set objControl   = Page.Controls.Item(x)
				Set objChildNode = GetSection(objControl.ControlID,node)
				If objControl.EnableViewState And Not objControl.IsViewStateRestored Then
					objControl.ReadProperties objChildNode
				End If
			Next
			Set objControl = Nothing
			Set objChildNode = Nothing
		End Function


		Public Function DataBind(ds, wc)			
			Dim x,mx
			Dim objControl
			
			If ds Is Nothing Then
				Exit Function
			End If

			mx = Page.Controls.Count - 1
			For x=0 To mx
				Set objControl   = Page.Controls.Item(x)
				If objControl.DataTextField <> "" Then
					On Error Resume Next						
					objControl.Owner.SetValueFromDataSource ds(objControl.DataTextField).Value
					If Err.number >0  Then
						Page.TraceError  objControl, "Bind Error",1
					End If
				End If
			Next			
			Set objControl = Nothing
		End Function
				
		Private Function PropagateOnInit()			
			Dim x,mx,objControl			
			mx = Page.Controls.Count - 1
			For x=0 To mx				
				Set objControl = Page.Controls.Item(x)
				If objControl.ImplementsOnInit Then
					Call objControl.Owner.OnInit()
				End If
			Next			
			Set objControl = Nothing
		End Function

		Private Function PropagateOnLoad()			
			Dim x,mx,objControl						
			mx = Page.Controls.Count - 1
			For x=0 To mx				
				Set objControl = Page.Controls.Item(x)
				If objControl.ImplementsOnLoad Then
					Call objControl.Owner.OnLoad()
				End If
			Next			
			Set objControl = Nothing
		End Function

		Private Function PropagateOnProcessPostBack()			
			Dim x,mx,objControl				
			mx = Page.Controls.Count - 1					
			For x=0 To mx							
				Set objControl = Page.Controls.Item(x)
				If objControl.ImplementsProcessPostBack Then					
					Call objControl.Owner.ProcessPostBack()
				End If
			Next			
			Set objControl = Nothing			
		End Function


		Public Function RegisterPostBackEventHandler(Control,EventHandlerName,Parameters)
			Dim x,obj
			If Not IsArray(mvarPostBackHandlers) Then
				Redim mvarPostBackHandlers(10,2) 'Maximum 10?
				mintPostBackHandlerIndex = 0
				For x = 0 To 10
					Set mvarPostBackHandlers(x,0) = Nothing
				Next
			End If
			
			If TypeName(Control) = "WebControl" Then
				Set obj = Control
			Else
				Set obj = Control.Control 'Gets the WebControl ...
			End If
			
			Set mvarPostBackHandlers(mintPostBackHandlerIndex ,0) = obj
				mvarPostBackHandlers(mintPostBackHandlerIndex ,1) = EventHandlerName
				mvarPostBackHandlers(mintPostBackHandlerIndex ,2) = Parameters			
			mintPostBackHandlerIndex = mintPostBackHandlerIndex + 1			
		End Function

		Private Function HandlePostBackHandlers()
			Dim x,mx
			Dim Fnc
			
			If IsArray(mvarPostBackHandlers) Then
				mx = UBound(mvarPostBackHandlers)
				For x = 0 To mx
					If mvarPostBackHandlers(x,0) Is Nothing Then
						Exit For
					Else
						Set Fnc = GetFunctionReference(mvarPostBackHandlers(x,0).ControlID & "_" & mvarPostBackHandlers(x,1))
						If Not Fnc Is Nothing Then
							Call Fnc(mvarPostBackHandlers(x,0),mvarPostBackHandlers(x,1))		
						End If
					End If
				Next	
			End If			
		End Function	

		Public Sub HandleServerEvent(EventName)
			Dim tStart
			tStart= Timer()		
		
			Page.TraceCall Page.Control, EventName	 & " - Start"
						
			Select Case EventName
				Case "Page_LoadViewState"									 
					 Call LoadViewState()
					 Call Page_ReadViewState()
					 IsViewStateLoaded = True
				Case "Page_SaveViewState"			
					Call Page_WriteViewState()
					IsViewStateSaved = True
				Case "Page_ProcessAjaxForm"
					Call Page_ProcessAjaxForm()
				Case "Page_ClientEvent"
					Response.Clear
					On Error Resume Next
					ExecuteEventFunctionParams Action,ActionSource
					If Err.number<> 0 Then
						Response.Clear
						Response.Write Err.number & ":" & Err.Description
						Response.End
					End If
					On Error Goto 0
					Page.HandleServerEvent "Page_SaveViewState"
					Call SaveViewState()
					Response.End
			End Select						
			
			Select Case EventName 
				Case  "RaiseChangedEvents"
					Call HandlePostBackHandlers()
				Case "ProcessPostBackData"
					Call PropagateOnProcessPostBack()
				Case Else
					ExecuteEventFunction EventName
			End Select
						
			Select Case EventName
				Case "Page_Init"										
					Call PropagateOnInit()
				Case "Page_Load"
					Call PropagateOnLoad()
			End Select
			
			Page.TraceCall Page.Control, EventName	 & " - End (" & FormatNumber( Timer()-tStart,6) & ")"

		End Sub

		Public Function GetPostBackEvent()
			Dim e
			If mPostBackEvent Is Nothing Then
				Set e = New ClientEvent
				e.EventName = Page.Action
				e.Source    = Page.ActionSource
				e.Instance  = Page.ActionSourceInstance
				e.ExtraMessage	= Page.ActionXMessage			
				Set mPostBackEvent = e
			End If
			Set GetPostBackEvent = mPostBackEvent
		End Function
		
		Public Function HandlePostBack(e)		
				Dim obj 
				Set e = GetPostBackEvent()
				HandlePostBack = False				
				If Control.ControlID = e.Source Then
					Set e.TargetObject  = Control
					Page.TraceImportantCall Page.Control, "<b>Page Event (" & e.EventName & ")</b>"
					HandlePostBack = ExecuteEventFunctionEX(e)
					Exit Function
				End If				
				
				Set obj = Controls.GetByName(e.Source)
				Set e.TargetObject  = obj
				
				If Not obj Is Nothing Then
					Page.TraceImportantCall obj, "Handled Event " & e.EventName & " : " & obj.HandleClientEvent(e)
					HandlePostBack = True
				Else
					Page.TraceCall Page.Control,"<span style='color:red;font-weight:bold'>Postback Without Action</span>"
					Page.TraceImportantCall Page.Control, e.Source & ":" & e.EventName
				End If							
				Set e = Nothing
		End Function
		
		Public Function ShowMessage(strTitle,strMessage,x,y,w,h)
				Dim strDivMsg
				strDivMsg = "<div id='__msg' class='thedummyclass' style='border:groove medium #0033cc;font-size:10pt;position: absolute; left: " & x & ";width:" & w & ";top:" & y & ";height:" & h & ";font-family: tahoma;background-color: #ffffcc;' onmouseover='___msg=this'>"
				strDivMsg = strDivMsg & "<table width='100%'><tr><td style='font-weight:bold;font-family:tahoma;color:blue'>" & strTitle & "</td><td align=right><a href=""""javascript:clasp.getObject('___msg').style.visibility ='hidden')"""">close</a></td></tr></table><hr>"
				strDivMsg= strDivMsg & strMessage
				strDivMsg = strDivMsg & "</div>"								
				Page.RegisterClientScript "showmessage",strDivMsg
		End Function
		
		Private Function GetIDFromControl(ctrl)
			If TypeName(ctrl) = "WebControl" Then
				GetIDFromControl = ctrl.ControlID
			Else
				GetIDFromControl = ctrl.Control.ControlID
			End If		
		End Function
				
		Public Function SetDefaultFocus(ctrl)
			Page.RegisterEventListener "body","onload","clasp.getObject(""" & GetIDFromControl(ctrl) & """).focus()"			
		End Function

		Public Function SetDefaultObject(ctrl,jsEvent)
			Call RegisterEventListener("document","onkeypress","clasp.getObject(""" & GetIDFromControl(ctrl) & """)." & jsEvent & "()") 
		End Function

		Public Function Execute()
			Call Main()
		End Function

		Public Sub RegisterWebControlsInClientSide()
			Dim x,mx,objControl						
			Dim strJs
			Dim ctrlID
			Set strJS = new StringBuilder
			mx = Page.Controls.Count - 1
			strJS.Append vbNewLine & "<script language='javascript'>" & vbNewLine
			For x = 0 To mx				
				ctrlID = Page.Controls.Item(x).ControlID
				strJS.Append  vbTab &  "clasp.page.addControl('" & ctrlID & "');" & vbNewLine				
			Next			
			Set objControl = Nothing
			strJS.Append "</script>" & vbNewLine			
			Response.Write strJs.ToString()
			Set strJs = Nothing
		End Sub
		
		Public Function OpenForm()		
			Call RenderClientImages() 'render registered images					
			Response.Write vbNewLine & vbNewLine
			Response.Write "<form name = 'frmForm' id='frmForm' action='" & PageFormAction & "' method='Post' onsubmit=""doPostBack('Submit','Page')"">"  & vbNewLine
			Response.Write "<input type='hidden' name='__FVERSION'   id='__FVERSION'   value='CJCA.YCMG.IACM.2.0'>" & vbNewLine
			Response.Write "<input type='hidden' name='__ACTION'     id='__ACTION'     value=''>" & vbNewLine
			Response.Write "<input type='hidden' name='__SOURCE'     id='__SOURCE'     value=''>" & vbNewLine
			Response.Write "<input type='hidden' name='__INSTANCE'   id='__INSTANCE'   value=''>" & vbNewLine
			Response.Write "<input type='hidden' name='__EXTRAMSG'   id='__EXTRAMSG'   value=''>" & vbNewLine	
			Response.Write "<input type='hidden' name='__ISPOSTBACK' id='__ISPOSTBACK' value='" & IsPostBack & "'>" & vbNewLine
			Response.Write "<input type='hidden' name='__ISREDIRECTEDPOSTBACK' id='__ISREDIRECTEDPOSTBACK' value=''>" & vbNewLine
			Response.Write "<input type='hidden' name='__SCROLLTOP'  id='__SCROLLTOP'  value='" & ScrollTop & "'>" & vbNewLine
			Response.Write "<input type='hidden' name='__SCROLLLEFT' id='__SCROLLLEFT' value='" & ScrollLeft & "'>" & vbNewLine
		End Function

		Public Function CloseForm()
			Call SaveViewState()
			Response.Write "<input type='hidden' name='__VIEWSTATECOMPRESSED'  id='__VIEWSTATECOMPRESSED'  value='" & Page.IsCompressed & "'>" & vbNewLine
			Response.write	"</form>" & vbNewLine & vbNewLine
			Response.Write(vbNewLine &  "<script language='JavaScript'>//CLASP -Misc Init Functions" & vbNewLine)
			If IsPostBack Then
				If AutoResetFocus Then
					Response.Write "clasp.events.addListener(document.body,'onload','clasp.events.setFocus(""" & ActionSource &  """," & ActionSourceInstance & ")');"					
				End If
				
				If AutoResetScrollPosition Then
					Response.Write "clasp.form.resetScroll();"
				End If
			End If

			If Page.OnSubmitStatement <> "" Then
				Response.Write vbNewLine &  "document.frmForm.__onSubmit=" & Page.OnSubmitStatement & ";"
			End If			

			Response.Write  vbNewLine &  "</script>" & vbNewLine 									
			Call RenderClientScripts()			
		End Function
		
		Public Function BackupViewState(bkName,ttl)	'ttl - time to live in minutes (not implemented)
			If Len(bkName &"") > 0 Then
				Page.HandleServerEvent "Page_SaveViewState" 'Collects the viewstate
				Session("VSBK__" & bkName) = CStr(Page.ScrollTop) & ":" & CStr(Page.ScrollLeft) & Page.ViewState.GetViewState()
			End If
		End Function

		Public Function RestoreViewState(bkName)
			Dim vbk
			Dim vsarr
			vbk = Session("VSBK__" & bkName)
			If Len(vbk)>0 Then
				vsarr = Split(vbk,":",3)
				IsPostBack = True   ' restores are treated as post backs
				Page.ScrollTop  = CLng(vsarr(0))
				Page.ScrollLeft = CLng(vsarr(1))				
				Call ViewState.LoadViewState(vsarr(2))
				Session("VSBK__" & bkName) = ""
			End If
		End Function

		Public Function OpenInFrame(Name,URL)
			Dim ScriptName
			ScriptName = "WindowFrame_" & Name
			RegisterClientScript ScriptName,"<script> clasp.openInFrame('"+Name+"','"+URL+"');</script>"
		End Function

	End Class


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::HELPER FUNCTIONS.

	Public Function GetSection(SectionName,Node)
		Set GetSection = GetSectionEx(SectionName,Node,"")		
	End Function
	
	Public Function GetSectionEx(SectionName,Node,IsNewSection)
		Dim obj			
		Set obj = Node.SelectSingleNode(SectionName)
		If  obj Is Nothing Then			
			Set GetSectionEx = Node.ownerDocument.CreateElement(SectionName)
			IsNewSection= True
		Else
			Set GetSectionEx = obj
			IsNewSection= False
		End If		
	End Function

	Public Function WriteProperty(Name,Value,Node)
		Dim att		
		Set att = Node.Attributes.getNamedItem(Name)
		If att Is Nothing Then
			Set att = Node.OwnerDocument.CreateAttribute(Name)
			Node.Attributes.SetNamedItem att
		End If
		att.Text = Value
		Set att = Nothing		
	End Function

	Public Function ReadProperty(Name,Node)		
		Dim obj
		Set obj =  Node.Attributes.GetNamedItem(Name)
		If obj Is Nothing  Then
			ReadProperty = ""
		Else
			ReadProperty = obj.Text
		End If
		Set obj = Nothing
	End Function

	Public Function ReadPropertyByIndex(Index,Node)		
		Dim obj
		Set obj =  Node.attributes.Item(Index)				
		If obj Is Nothing  Then
			ReadPropertyByIndex = ""
		Else
			ReadPropertyByIndex = obj.Text
		End If		
		Set obj = Nothing
	End Function
	
	Public Function ExecuteEventFunction(EventName)
		Dim fnc
		ExecuteEventFunction = False		
		Set fnc = GetFunctionReference(EventName)		
		If Not fnc Is Nothing Then
			ExecuteEventFunction = True			
			Call fnc()
		End If
	End Function	
	
	Public Function ExecuteEventFunctionWParam(EventName,Params)
		Dim fnc
		ExecuteEventFunctionParam = False		
		Set fnc = GetFunctionReference(EventName)		
		If Not fnc Is Nothing Then
			ExecuteEventFunctionParam = True			
			Call fnc(Params)
		End If	
	End Function
	
	Public Function ExecuteEventFunctionEX(e)
		Dim fnc
		ExecuteEventFunctionEX = False		
		Set fnc = GetFunctionReference(e.EventFnc)		
		If Not fnc Is Nothing Then
			ExecuteEventFunctionEX = True			
			Call fnc(e)
		End If
	End Function

	Public Function ExecuteEventFunctionParams(EventName,p)
		Dim fnc		
		ExecuteEventFunctionParams = False
		Set fnc = GetFunctionReference(EventName)		
		If Not fnc Is Nothing Then
			ExecuteEventFunctionParams = True			
			Call fnc(p)
		End If
	End Function	
	
	Public Function GetFunctionReference(sFncName)
		On Error Resume Next
			Set GetFunctionReference = Nothing
			Set GetFunctionReference = GetRef(sFncName)		
		On error Goto 0
	End Function
	
	Public Function IIf(expression, truecondition,falsecondition)
		If CBool(expression) Then
			IIf = truecondition
		Else
			IIf = falsecondition
		End If
	End Function

	Function URLDecode(str)
		dim re
		set re = new RegExp

		str = Replace(str, "+", " ")
		
		re.Pattern = "%([0-9a-fA-F]{2})"
		re.Global = True
		URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
	End Function

	' Replacement function for the above
	Function URLDecodeHex(match, hex_digits, pos, source)
		URLDecodeHex = chr("&H" & hex_digits)
	End function

%>
