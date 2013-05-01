<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Button

	'Helper function.
	Public Function New_ServerButton(name) 
		Set New_ServerButton = New ServerButton
			New_ServerButton.Control.Name = name
			New_ServerButton.Mode = 1
	End Function

	Public Function New_ServerLinkButton(name) 
		Set New_ServerLinkButton = New ServerButton
			New_ServerLinkButton.Control.Name = name
			New_ServerLinkButton.Mode = 2
	End Function

	Public Function New_ServerImageButton(name) 
		Set New_ServerImageButton = New ServerButton
			New_ServerImageButton.Control.Name = name
			New_ServerImageButton.Mode  = 3
	End Function

	Public Function New_ServerAdvanceButton(name) 
		Set New_ServerAdvanceButton = New ServerButton
			New_ServerAdvanceButton.Control.Name = name
			New_ServerAdvanceButton.Mode = 4
	End Function


	Page.RegisterLibrary "ServerButton"

	 Class ServerButton
		Dim Control
		Dim Text
		Dim OnClick
		Dim CommandSource
		Dim Mode '1 Button, 2 Link Button, 3 Image Button, 4 Advance Button
		
		Dim CausesValidation
		Dim ValidationGroup
		
		'Img		
		Dim Image
		Dim RollOverImage
		Dim PostMode '0 Normal, 1 Ajax

		'Link Button
		Dim Target
		
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			OnClick = ""
			Text    = ""				
			Mode    = 1
			Image   = ""
			RollOverImage = ""
			CausesValidation = True
			ValidationGroup  = ""
			PostMode		 = 0
	   End Sub
	   
	   
	   	Public Function ReadProperties(bag)			
			Text  = bag.Read("T")
			Image = bag.Read("I")
			RollOverImage = bag.Read("RI")
			Target = bag.Read("TG")			
			Mode   = CInt(bag.Read("M"))
			CausesValidation = CBool(bag.Read("V"))
			ValidationGroup  = bag.Read("G")
			PostMode = bag.Read("P")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",  Text
			bag.Write "I",  Image
			bag.Write "RI",RollOverImage
			bag.Write "TG", Target	
			bag.Write "M",  Mode
			bag.Write "V",  CausesValidation
			bag.Write "G",  ValidationGroup
			bag.Write "P",  PostMode
		End Function
	   
	   Public Function HandleClientEvent(e)
		 	If OnClick<>"" Then
				HandleClientEvent = ExecuteEventFunction(OnClick)
			Else			
				HandleClientEvent = ExecuteEventFunction(e.EventFnc)
		 	End If	
	   End Function			
	   
	   	Public Function SetValueFromDataSource(value)
			If Mode = 3 Then
				Image = value
			Else
				Text = value
			End If
	    End Function

	   Private Function RenderButton(EventHandler)
			Response.Write "<input type='Button' id='" & Control.ControlID & "' name='"  & Control.ControlID & "' "   &_
				" value=""" & Server.HTMLEncode(Text) & """ " &_
				" tabindex= " & Control.TabIndex &_
				EventHandler
				
				If Control.Style<>"" Then Response.Write " style='" & Control.Style  & "' "
				If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
				If Target<>"" Then Response.Write " Target='" & target + "' "
				If Not Control.Enabled Then Response.Write " disabled "
				If Not CausesValidation Then Response.Write " skipvalidation=true "
				If ValidationGroup<>"" Then Response.Write " validationgroup='" & ValidationGroup + "' "
				    				
			 Response.Write ">" & vbNewLine
	   End Function

	   Private Function RenderLinkButton(EventHandler)
			Response.Write "<a id='" & Control.ControlID & "'" &_
						EventHandler
			If Control.Style<>"" Then Response.Write " style='" & Control.Style  & "' "
			If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
			If Target<>"" Then Response.Write " target='" & Target + "' "
			If Not CausesValidation Then Response.Write " skipvalidation=True "
			If ValidationGroup<>"" Then Response.Write " validationgroup='" & ValidationGroup + "' "
			Response.Write ">"
			If Image <> "" Then
				Response.Write "<img src='" & Image & "' border=0>"
			End If
			Response.Write Text & "</a>"  & vbNewLine	   
	   End Function

	   Private Function RenderImageButton(EventHandler)
	   		
	   		Dim RollOverScript
			If RollOverImage <> "" Then			 
				RollOverScript = " onmouseover = 'this.src =" & chr(34) & RollOverImage & chr(34) & "' " & _
							     " onmouseout  = 'this.src =" & chr(34) & Image & chr(34) & "' "
			End If							    

			Response.Write "<a " & EventHandler & " id='" & Control.ControlID & "'>"		 
			Response.Write "<img border=0 id='" & Control.ControlID & "_img' " &_ 
					       " tabindex = " & Control.TabIndex &_
					       " src = '" &  Image & "'" &_
					       RollOverScript			
			If Control.Style<>"" Then Response.Write " style='" & Control.Style  & "' "
			If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
			If Target<>"" Then Response.Write " target='" & Target + "' "
			If Not CausesValidation Then Response.Write " skipvalidation=True "
			If ValidationGroup<>"" Then Response.Write " validationGroup='" & ValidationGroup + "' "
			Response.Write "></a>" & vbNewLine
	   
	   End Function

	   Private Function RenderAdvanceButton(EventHandler)
			Response.Write "<button id='" & Control.ControlID & "' name='"  & Control.ControlID & "' "   &_
				" tabindex= " & Control.TabIndex  
				If Control.Style<>""	Then Response.Write " style='" & Control.Style  + "' "
				If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
				If Target<>""			Then Response.Write " target='" & Target + "' "
				If Not Control.Enabled	Then Response.Write " disabled "
				If Not CausesValidation Then Response.Write " skipvalidation=true "
				If ValidationGroup<>"" Then Response.Write " validationgroup='" & ValidationGroup + "' "
				Response.Write EventHandler & ">" 
				Response.Write  Text
				Response.Write "</button>"					
	   End Function

	   
	   Public Default Function Render()
			
			Dim varStart	 
			Dim evtName
			Dim JsEvent

			If Control.Visible = False Then
				Exit Function
			End If

			varStart = Now

			If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			End If

			If Mode = 1 or Mode = 4 Then
				JsEvent = "onclick"
			Else
				JsEvent = "href"
			End If

			'To support the ItemCommand
			Select Case TypeName(Control.Parent.Owner) 
				Case  "ServerDBTable" , "ServerDataRepeater", "ServerDataList" 
					If CommandSource = "" Then					
						evtName = Page.GetEventScript(JsEvent,Control.ControlID,"ItemCommand",0,Control.Parent.Owner.DataSource.AbsolutePosition)
					Else
						evtName = Page.GetEventScript(JsEvent,Control.ControlID,"ItemCommand",0,CommandSource)
					End If					
				Case Else	
					If PostMode = 0 Then
						evtName = Page.GetEventScript(JsEvent,Control.ControlID,"OnClick",0,CommandSource)
					Else
						evtName = Page.GetAjaxEventScript(Me,Control.ControlID & "_OnServerEvent",false,JsEvent)
					End If
			End Select			 
	
			Select Case Mode
				Case 1:RenderButton evtName
				Case 2:RenderLinkButton evtName
				Case 3:RenderImageButton evtName
				Case 4:RenderAdvanceButton evtName
				Case Else
					RenderButton evtName
			End Select
			Page.TraceRender varStart,Now,Control.ControlID
		End Function

	End Class

%>