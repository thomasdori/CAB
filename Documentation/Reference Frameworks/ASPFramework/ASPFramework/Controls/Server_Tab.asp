<Script language="JavaScript">
	var __tabhack="";
	function __SetTabStyle(tab,style) {
			var ci = document.frmForm.MainTab_I
			if(ci) {				
				if(ci.value != tab.Index)	{
					if(event.type == 'mouseover' || __tabhack=='over') tab.style.cssText = tab.MOStyle;	
					if(event.type=='mouseout' || __tabhack=='out')    tab.style.cssText = tab.DEStyle;
				}else {
					if(event.type == 'mouseover'  || __tabhack=='over') tab.style.cssText = tab.SMOStyle;	
					if(event.type=='mouseout' || __tabhack=='out')    tab.style.cssText = tab.SDEStyle;
				}

			}
			else {
				tab.style.cssText = style;				
			}
			__tabhack="";
	}//
	
	function __ShowTab(tab,base) {
	 
			var obj = eval(tab.id + "_contents")
			var x,mx;
			eval("document.frmForm." + base + "_I").value = -1
			for(x=0;x<100;x++) {				
				try {
					var test = eval(base + "_" + x + "_contents");
					var testTab = eval(base + "_" + x);
															
					if(test.style.visibility!='hidden') {
						 __tabhack="out"
						 testTab.onmouseout();						
						 test.style.visibility = 'hidden';
					}
					if(testTab == tab) {						
						eval("document.frmForm." + base + "_I").value = testTab.Index;						
					}else  {
					}
					
				}//
				catch(e) { break; }
			}		
			__tabhack="over"				
			tab.onmouseover();				
			obj.style.visibility = "visible";
	}


</script>

<%

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Tab

	Public Function New_ServerTabStrip(name) 
		Set New_ServerTabStrip = New ServerTabStrip
			New_ServerTabStrip.Control.Name = name
	End Function

	Page.RegisterLibrary "ServerTab"

	Class ServerTab
		
		Dim SelectedStyle
		Dim SelectedHoverStyle
		
		Dim Style
		Dim HoverStyle
														
		Dim RenderFunction
		Dim Caption

		Dim ID
		Dim Owner
		Dim TabIndex
				

		Public Function RenderCaption(Width,ColSpan,IsVisible)
		
			Dim DefaultStyle
			Dim MouseOverStyle
			Dim MouseOutStyle
			
			If IsVisible Then			
				DefaultStyle   = "cursor:hand;" & IIF(SelectedStyle<>"",SelectedStyle,Style)
				MouseOverStyle = IIF(SelectedHoverStyle<>"","cursor:hand;" & SelectedHoverStyle,DefaultStyle)
			Else
				DefaultStyle   = "cursor:hand;" & Style
				MouseOverStyle = IIF(HoverStyle<>"","cursor:hand;" & HoverStyle,DefaultStyle)
			End If
			
			MouseOutStyle  = DefaultStyle			
			
			Response.Write vbNewLine &  "<TD "			
			Response.Write " onmouseover='__SetTabStyle(this,""" & MouseOverStyle & """)'"
			Response.Write " onmouseout='__SetTabStyle(this,""" & MouseOutStyle & """)'"
			Response.Write " style='" & DefaultStyle & "'"
			Response.Write " id='" & ID & "' " 
			
			If Owner.AutoPostBack Then
				Response.Write Page.GetEventScript("onclick", Owner.Control.ControlID, "TabClick", TabIndex,"")
			Else
				Response.Write " onclick = ""__ShowTab(this,'" & Owner.Control.ControlID & "')"" "
			End If
  	
			Response.Write IIF(Width<>""," Width='" & Width & "' ","")			
			Response.Write IIF(ColSpan<>""," ColsPan='" & ColSpan & "' ","")
			Response.Write "><SPAN>" & Caption
			Response.Write "</SPAN></TD>"  &  vbNewLine
			If Not Owner.AutoPostBack Then
			Dim script
			script = "<script language='javascript'>{ var v=document.getElementById('" & ID  & "');" & vbNewLine &_
										 "v.DEStyle='" & "cursor:hand;" & Style & "';" &  vbNewLine &_ 
										 "v.MOStyle='" & IIF(HoverStyle<>"","cursor:hand;" & HoverStyle,DefaultStyle) & "'"  & vbNewLine &_
										 "v.SDEStyle='" & "cursor:hand;" & IIF(SelectedStyle<>"",SelectedStyle,Style) & "'" & vbNewLine  &_
										 "v.SMOStyle='" & IIF(SelectedHoverStyle<>"","cursor:hand;" & SelectedHoverStyle,DefaultStyle) & "'" & vbNewLine &_
										 "v.Index=" & TabIndex  & vbNewLine &_
									     "}</script>"
			Page.RegisterClientScript ID,script
			End If
		End Function
		
		Public Function Render(IsVisible)
			Dim fnc
			Dim MyStyle
			
			Set fnc = Nothing				
			If IsVisible Then
				MyStyle = "top:0px:left:0px;width:100%;height:100%; overflow: auto; position: absolute;"
			Else
				MyStyle = "visibility:hidden;top:0px:left:0px;width:100%;height:100%; overflow: auto; position: absolute;"
			End If
			
			Response.Write vbNewLine &  "<SPAN id='" & ID & "_contents' style='" & MyStyle & "'>"
						
			Set fnc = GetFunctionReference(RenderFunction)
			
			If Not fnc Is Nothing	Then
				Call fnc()
			End If
			
			Response.Write "</SPAN>" & vbNewLine
			
		End Function
	
	End Class


	 Class ServerTabStrip		
		'Public Controls
		Public Control
		
		Public TabsPerRow
		Public TabsWidth		
		Public Tabs

		Public Orientation 'Vertical, Horizontal (DEF)
		Public Width
		Public Height
		Public BorderWidth
		Public BorderColor
		
		Public SelectedTabIndex
		Public AutoPostBack
		
		Public Style
		Public CssName
				
		Public DefaultStyle
		Public DefaultHoverStyle
		
		Public DefaultSelectedStyle
		Public DefaultSelectedHoverStyle
						
		Public TabStripTabSeparatorTemplate
		Public TabSeparatorTemplate
		
						
		Private Sub Class_Initialize()
			Set Control = New WebControl	
		'	Set Controls = Control.Controls
			Set Control.Owner = Me 			
			TabsPerRow = -1
			SelectedTabIndex = 0
			AutoPostBack = True
	    End Sub
	   
	   	Public Function ReadProperties(bag)				   			   		
			If Not AutoPostBack Then
				 SelectedTabIndex = CInt(Request.Form(Control.ControlID & "_I"))
			Else
	   			SelectedTabIndex = CInt(bag.Read("STI"))
	   		End If
	   	End Function
		
		Public Function WriteProperties(bag)
			bag.Write "STI",SelectedTabIndex
		End Function

	   
	    Public Function HandleClientEvent(e)
			SelectedTabIndex = CInt(e.Instance)
	    End Function			
	   
		Private Function RenderTab()

			Dim objTab
			Dim Row,Rows
			Dim Col,Cols
			Dim x
			Dim LastRowWidthModifier
				
			If Not IsArray(Tabs) Then
				Response.Write "ERROR!"
				Exit Function
			End If
			
			If TabsPerRow < 1 Then
				TabsPerRow = Ubound(Tabs)+1
			End If
			
			Rows = CInt(( Ubound(Tabs) + 1 ) / TabsPerRow)
			Cols  = TabsPerRow
			LastRowWidthModifier = ( Ubound(Tabs) + 1 ) Mod TabsPerRow
				
			Response.Write "<TABLE Border=0 CellSpacing=0  CellPadding=5px "
			Response.Write IIF(TabsWidth<>"", " Width=" & TabsWidth,"")
			Response.Write ">"
			x = 0
			
			If Rows = 0 And Ubound(Tabs)>0 Then
				Rows = 1
			End If
			
			For Row = 1 To Rows
				Response.Write "<TR>"																
				For Col = 1 To  IIF(Row = 1,LastRowWidthModifier,Cols)
					 Set objTab = Tabs(x)
					 Set objTab.Owner = Me
					 objTab.TabIndex = x
					 objTab.ID = Control.ControlID & "_" & x 
					 Call objTab.RenderCaption("","",(SelectedTabIndex=x))
					 x = x + 1
				Next
				Response.Write "</TR>"	
			Next
			Response.Write "</TABLE>"
			
			If TabStripTabSeparatorTemplate<>"" Then
				ExecuteEventFunction TabStripTabSeparatorTemplate
			End If

			Response.Write "<TABLE CellPadding=0 CellSpacing=0"
			If BorderWidth > 0 Then
				Response.Write " Border=" & BorderWidth & " BorderColor='" & BorderColor & "'"
			End If
			
			Response.Write IIF(Width<>"", " Width=" & Width ,"")
			Response.Write IIF(Style<>"", " Style=" & Style ,"")
			Response.Write IIF(CssName<>"", " Class=" & CssName ,"")
			Response.Write ">"
			Response.Write "<TR "
			Response.Write IIF(Height<>"", " Height=" & Height ,"") & ">"
			Response.Write "<TD>"			
			x=0
			
				For Each objTab in Tabs
					If AutoPostBack Then
						If SelectedTabIndex=x Then
							Call objTab.Render((SelectedTabIndex=x))
						End If
					Else
						Call objTab.Render((SelectedTabIndex=x))
					End If
					x=x+1
				Next				
			Response.Write "</TD></TR></TABLE>"
			'keep track of it...
			If Not AutoPostBack Then
				Response.Write "<input type=HIDDEN id='" & Control.ControlID & "_I' name='" & Control.ControlID & "_I' Value='" & SelectedTabIndex & "'>"
			End If
		End Function
				
	    Public Default Function Render()
				Call RenderTab()
	    End Function
	  
	End Class

%>