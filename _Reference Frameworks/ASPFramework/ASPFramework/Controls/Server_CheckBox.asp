<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server TextBox

	Dim CSS_Name_ServerCheckBox
	Dim CSS_Name_ServerCheckBox_Caption
	CSS_Name_ServerCheckBox_Caption = ""
	
	Public Function New_ServerCheckBox(name) 
		Set New_ServerCheckBox = New ServerCheckBox
			New_ServerCheckBox.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerCheckBox"
	
	 Class ServerCheckBox
		Dim Control
		Dim Caption
		Dim ReadOnly
		Dim Checked
		Dim RaiseOnChanged
		Dim AutoPostBack
		Dim OnClick
						
		Private mbolWasRendered
			
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Control.ImplementsProcessPostBack = True
			Control.SupportsClientSideEvent   = True			
			
			Caption        = ""
			ReadOnly       = False
			Checked        = False
		    AutoPostBack   = False
		    RaiseOnChanged = False
			OnClick = ""						
			mbolWasRendered = False
	    End Sub
	   	
	   	Public Function WriteProperties(bag)			
			 bag.Write "R",False
			 bag.Write "P",AutoPostBack
			 bag.Write "C",Caption
			 bag.Write "E",OnClick
			 bag.Write "U",RaiseOnChanged 
			 
			 If Control.Visible Then										
			  	bag.Write "S", Checked
			End If
		End Function
		
		Public Function ReadProperties(bag)
			Caption         = bag.Read("C")
			OnClick         = bag.Read("E")
			AutoPostBack    = CBool(bag.Read("P"))
			RaiseOnChanged  = CBool(bag.Read("U"))			
			mbolWasRendered = CBool(bag.Read("R"))
			Checked		    = CBool(bag.Read("S"))		
		End Function

		Public Function ProcessPostBack()
			
			'If ViewState is present AND control was not rendered or was disabled then don't process.
			If Not Control.ViewState Is Nothing And Not (mbolWasRendered AND Control.Enabled)  Then
				Exit Function
			End If				
			
			Checked  = (Request.Form(Control.ControlID)<>"")
			If RaiseOnChanged  Then	
				If Checked  <> CBool( Control.ViewState.Read("S")) Then
					Call Page.RegisterPostBackEventHandler(Me,"OnChanged","")
				End If
			End If

		End Function

	   
	   Public Function HandleClientEvent(e)
		 	If AutoPostBack Then
		 		
		 		If OnClick <> "" Then		 			
		 			HandleClientEvent = ExecuteEventFunction(OnClick)
		 		Else
		 			HandleClientEvent = ExecuteEventFunction(e.EventFnc)
		 		End If
		 	End If	
	   End Function			
	   
	   Public Function SetValueFromDataSource(value)
			Checked = CBool(value)			
	   End Function

	   Public Function SetValue(value) 
			Checked = CBool(value)
	   End Function
		
	   Private Function RenderCheckBox()
			Dim style
			Dim evtName
	   		
	   		If AutoPostBack Then
	   			evtName = Page.GetEventScript("onclick",Control.ControlID,"Click","","")
	   		End If
	   		
			Response.Write "<input type='checkbox' id='" & Control.ControlID & "' name='" & Control.ControlID & "' " &_
					          iif(Checked," checked "," ") &_
							  " tabindex = " & Control.TabIndex &_
					          " value = '1' "  &_
					          IIf(Control.Style<>""," style='" & Control.Style + "' ","") &_
						      IIf(Not Control.Enabled," disabled ","") &_
						      IIf(Control.CssClass<>""," class='" & Control.CssClass  + "' ","") &_
						      evtName  &_
					         ">&nbsp;"	& IIF(Caption<>"","<span class='" & CSS_Name_ServerCheckBox_Caption & "'>" &  Caption & "</span>","") 
	   End Function

	   Public Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.IsVisible = False Then
				Exit Function
			 End If

			 varStart = Now
 
 			 If Not Control.ViewState Is Nothing Then 				
				Control.ViewState.Write "R",True
 			 End If
			 
			 If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			 End If
			 
			 Render = RenderCheckBox()
					
		 	Page.TraceRender varStart,Now,Control.ControlID
		 	 
		End Function

	End Class

%>