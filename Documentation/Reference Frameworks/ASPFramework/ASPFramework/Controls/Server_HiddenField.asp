<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Hidden Field
	
	'Helper function.
	Public Function New_ServerHiddenField(name) 
		Set New_ServerHiddenField = New ServerHiddenField
		New_ServerHiddenField.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerHiddenField"
	
	 Class ServerHiddenField
		Dim Control
		Dim Value
		Dim ValueChanged
		Dim RaiseOnChanged
		
		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Control.ImplementsProcessPostBack = True			
			Control.SupportsClientSideEvent   = True

			Value    = ""	
			ValueChanged   = False
			RaiseOnChanged = False
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			RaiseOnChanged = CBool(bag.Read("RC"))
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "V", Value
			bag.Write "RC",RaiseOnChanged 
		End Function
	   
		Public Function ProcessPostBack()
			If  Request.Form.Key(Control.ControlID) <> "" Then
				If Not Control.ViewState Is Nothing Then
					Value        = Request.Form(Control.ControlID)
					ValueChanged = (Control.ViewState.Read("V") <> Value)
					If RaiseOnChanged Then
						Call Page.RegisterPostBackEventHandler(Me,"OnChanged","")
					End If
				End If
			Else
				If Not Control.ViewState Is Nothing Then
					Value  = Control.ViewState.Read("V")
				End If
			End If
		
		End Function

		Public Function SetValue(value) 
			Value = value
		End Function
		
	   Public Function HandleClientEvent(e)
			'Nothing, right?
	   End Function			
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			
			If Control.Visible = False Then
				Exit Function
			End If			 

			varStart = Now

			Response.Write "<input type='hidden' id='" & Control.ControlID  & "' name = '" & Control.ControlID & "' value='" & Server.HTMLEncode(Value) & "'>"
			Page.TraceRender varStart,Now,Control.ControlID

		End Function

	End Class

%>