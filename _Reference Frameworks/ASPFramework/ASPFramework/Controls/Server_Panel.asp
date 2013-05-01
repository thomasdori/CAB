<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Panel
	
	Public Function New_ServerPanel(name,width,height) 
		Set New_ServerPanel = New ServerPanel
			New_ServerPanel.Control.ControlID = name
			New_ServerPanel.Width = width
			New_ServerPanel.Height = height
	End Function

	Page.RegisterLibrary "ServerPanel"

	Page.RegisterClientStartupScript "PanelScripts", _
		"<script language=""JavaScript"">" + _
		"	function Panel_DoScroll(pnlID,scrollPos) {" + _
		"		var elm;" + _
		"		var o,d,t,l;" + _
		"		d = clasp.getObject(pnlID+'__scroll');" + _
		"		if (!d) return null;" + _
		"		if (scrollPos) { " + _
		"			var a = scrollPos.split('.');" + _
		"			d.scrollTop = a[0];" + _
		"			d.scrollLeft = a[1];" + _
		"		}" + _
		"		t = parseInt(d.scrollTop);" + _
		"		l = parseInt(d.scrollLeft);" + _
		"		var elm = clasp.getObject(pnlID+'__scrollElm');" + _
		"		if(elm) elm.value = t+'.'+l;" + _
		"	};" + _
		"</script>"
	
	Class ServerPanel
	
		Dim Control
		Dim Text
		Dim PanelTemplate		
	
		Dim Width
		Dim Height
		Dim Mode		' 1 =  FlatPanel, 2 = 3DPanel, 3 = RoundPanel
		Dim BorderColor
		Dim LightColor
		Dim DarkColor
		Dim BackColor
		Dim AutoSize
		Dim EnableScrollBars

		Private mContents() 	' WebControl Objects (and string, etc) to be rendered inside the layer			
		Private mScrollPos 'stores  scroll position

				
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 
			Control.ImplementsProcessPostBack = True			
			
			Width		= 0
			Height		= 0
			Mode 		= 1 'flat mode
			Text		= ""
			BorderColor	= "#E0E0E0"
			LightColor	= "#FFFFFF"
			DarkColor	= "#C0C0C0"
			BackColor	= "#EFEBD7"
			AutoSize	= False
			PanelTemplate = ""
			EnableScrollBars = False

			ReDim mContents(0)
			
			Call PreloadImages() 'preload images
		End Sub
	   
	   	Public Function ReadProperties(bag)
			Mode			= bag.Read("M")
			EnableScrollBars= bag.Read("S")
			Text			= bag.Read("T")
			Width			= bag.Read("W")
			Height			= bag.Read("H")
			AutoSize		= bag.Read("A")
			PanelTemplate	= bag.Read("P")
			mScrollPos		= bag.Read("V")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "M",Mode
			bag.Write "S",EnableScrollBars
			bag.Write "T",Text
			bag.Write "W",Width
			bag.Write "H",Height
			bag.Write "A",AutoSize
			bag.Write "P",PanelTemplate
			bag.Write "V",mScrollPos
		End Function

		Public Function ProcessPostBack()
			Dim key:key = Control.ControlID &"__scrollElm"	
			If Request.Form(key) <> "" Then
				mScrollPos = Request.Form(key)
			End If		
		End Function
		
		Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunctionEX(e)
		End Function

	   	Public Function AddContent(ByRef Cn)
	   		Dim ub
	   		
	   		ub = UBound(mContents)
	   		
	   		If IsObject(cn) Then 
				' Check for WebControl
				On Error Resume Next
				If Cn.Control is Nothing Then Exit Function
				If Err.Number<>0 then Exit Function
				Set mContents(ub) = Cn	' only WebControl is to be used as an object
				'Set Cn.Control.Owner = Me	' ok to generate error if Cn is not a WebControl
			Else
				mContents(ub) = Cn		'other datatypes - string, number, etc
			End If	   		
	   		ReDim Preserve mContents(ub+1)
	   	End Function	   
	   
	   
		Private Function OpenRenderTag()		
			Dim Style
			Dim PanelStyle
			
			
			Style = Control.Style
			PanelStyle = " style='cursor:default; " + _
						IIf(Mode=1 And Not AutoSize,"width:"& Width &"; height:"& Height &"; ","") + _
						IIf(EnableScrollBars<>"","overflow:auto;clip:auto; ","") +"'"
			
			Response.Write vbNewLine & "<div  id='" & Control.ControlID  & "'" & _
				   IIf(Style<>""," Style='" & Style  + "' ",PanelStyle) & _
				   IIf(Control.CssClass<>"","Class='" & Control.CssClass  + "' ","") & _
				   IIf(Not Control.Enabled," disabled='disabled' ","") & _
				 ">" 
			If EnableScrollBars Then Response.Write "<input type='hidden' id='"+Control.ControlID+"__scrollElm' name='"+Control.ControlID+"__scrollElm' value='"& mScrollPos &"'>" & vbNewLine
			OpenRenderTag = True

		End Function
		
		Private Function CloseRenderTag()
			Response.Write "</div>"
			If EnableScrollBars Then 
				Response.Write "<script>Panel_DoScroll('"& Control.ControlID &"','"& mScrollPos &"')</script>"
			End If
		End Function
		
		Public Default Function Render ()
			 Dim varStart	 
			 
			 varStart = Now
			 
			 If Control.Visible = False Then
				Exit Function
			 End If

			Call OpenRenderTag()
				Select Case Mode
					Case 2 		'3DPanel
						Call Render3DPanel
					Case 3		'RoundPanel
						Call RenderRoundPanel
					Case Else	'flat
						Call RenderFlatPanel
				End Select
			Call CloseRenderTag()
			
			Page.TraceRender varStart,Now,Control.ControlID
		End Function
		
		Private Function RenderHTMLContent()
			Dim i
			Dim fnc
			If PanelTemplate = "" Then
				If Text <> "" Then
					Response.Write Text
				End If
			Else
				On Error Resume Next
				If Me.Control.Parent.Name = "Page" Then
					Set fnc = GetRef(PanelTemplate)	
					On Error Goto 0
					If Not fnc Is Nothing Then
						Call fnc()
					End If
				Else
					Me.Control.Parent.Owner.RenderHTMLContent()
				End If
			End If
			'content array
			For i=0 to UBound(mContents)-1 ' last item is always empty
				Response.Write mContents(i) & ""
			Next						
		End Function

		Private Function RenderFlatPanel()
			Dim scroll
			If Len(EnableScrollBars)>0 Then scroll = "overflow:auto" 
		%>
			<div id="<%=Control.ControlID%>__scroll"  style= "width:<%=Width%>; height:<%=Height%>;<%=scroll%>" 
				<%=IIf(Not Control.Enabled," Disabled ","")%>
				<%=IIf(EnableScrollBars,"onscroll=""Panel_DoScroll('"& Control.ControlID &"')"" ","")%>
				>
				<%RenderHTMLContent%>
			</div>
		<%
		End Function
		
		Private Function Render3DPanel()
			Dim scroll
			If Len(EnableScrollBars)>0 Then scroll = "overflow:auto" 
		%>
			<table border="0" width="<%=Width%>" height="<%=Height%>" cellspacing="0" cellpadding="0" style="border:1px solid <%=BorderColor%>">
			<tr>
				<td width="100%" style="border-left:1px solid <%=LightColor%>; border-right: 1px solid <%=DarkColor%>; border-top: 1px solid <%=LightColor%>; border-bottom: 1px solid <%=DarkColor%>" bgcolor="<%=BackColor%>">
				<div id="<%=Control.ControlID%>__scroll"  style= "width:<%=Width%>; height:<%=Height%>;<%=scroll%>" 
					<%=IIf(Not Control.Enabled," Disabled ","")%>
					<%=IIf(EnableScrollBars,"onscroll=""Panel_DoScroll('"& Control.ControlID &"')"" ","")%>
					>
					<%RenderHTMLContent%>
				</div>
				</td>
			</tr>
			</table>
		<%
		End Function

		Private Function RenderRoundPanel
			Dim pth
			Dim scroll

			If Len(EnableScrollBars)>0 Then scroll = "overflow:auto"
			pth = SCRIPT_LIBRARY_PATH
			If Width<30 Then Width = 30
			If Height<30 Then Height = 30
		%>
			<table border="0" cellspacing="0" cellpadding="0"><tr>
			<td valign="top" nowrap><img border="0" src="<%=pth%>panel/images/pnltl.gif" width="15" height="15"></td>
			<td valign="top" align="right" background="<%=pth%>panel/images/pnlt.gif" width="<%=Width-30%>"></td>
			<td valign="top" align="right" nowrap><img border="0" src="<%=pth%>panel/images/pnltr.gif" width="15" height="15"></td>
			</tr><tr>
			<td valign="bottom" background="<%=pth%>panel/images/pnll.gif" height="<%=Height-30%>"></td>
			<td>
			<div id="<%=Control.ControlID%>__scroll" style= "width:<%=Width-30%>; height:<%=Height-30%>;<%=scroll%>" 
				<%=IIf(Not Control.Enabled," Disabled ","")%>
				<%=IIf(EnableScrollBars,"onscroll=""Panel_DoScroll('"& Control.ControlID &"')"" ","")%>
				>
				<%RenderHTMLContent%>
			</div>
			</td>
			<td valign="bottom" align="right" background="<%=pth%>panel/images/pnlr.gif"></td>
			</tr><tr>
			<td valign="bottom"><img border="0" src="<%=pth%>panel/images/pnlbl.gif" width="15" height="15"></td>
			<td valign="bottom" align="right" background="<%=pth%>panel/images/pnlb.gif" height="15"></td>
			<td valign="bottom" align="right"><img border="0" src="<%=pth%>panel/images/pnlbr.gif" width="15" height="15"></td>
			</tr></table>
		<%
		End Function
		
		Private Function PreloadImages()
			Dim pth
			pth = SCRIPT_LIBRARY_PATH
			Page.RegisterClientImage "pnltl",pth+"panel/images/pnltl.gif"
			Page.RegisterClientImage "pnlt",pth+"panel/images/pnlt.gif"
			Page.RegisterClientImage "pnltr",pth+"panel/images/pnltr.gif"
			Page.RegisterClientImage "pnlr",pth+"panel/images/pnlr.gif"
			Page.RegisterClientImage "pnll",pth+"panel/images/pnll.gif"
			Page.RegisterClientImage "pnlbl",pth+"panel/images/pnlbl.gif"
			Page.RegisterClientImage "pnlb",pth+"panel/images/pnlb.gif"
			Page.RegisterClientImage "pnlbr",pth+"panel/images/pnlbr.gif"
		End Function
		
	End Class
%>
