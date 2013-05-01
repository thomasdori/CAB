<%
	Public Function New_ServerForm(name) 
		Set New_ServerForm = New ServerForm
			New_ServerForm.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerForm"
	
	Class ServerForm
		Dim Control
		Dim Controls
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			'Only if control is a container... if not don't expose it! (they can get to it anyway)
			Set Controls = New ControlsCollection
			
			Control.ImplementsOnInit = True '?  If you want the OnInit to be called
			Control.ImplementsOnLoad = True '?  If you want the Onload to be called
			Control.ImplementsProcessPostBack = True 'If you want to receive a ProcessPostBack notification on postbacks for form inputs
			
	   End Sub
	   
	   	Public Sub OnInit()			
		End Sub

	   	Public Sub OnLoad()			
		End Sub
		
	   Public Function SetValueFromDataSource(value)
			
	   End Function
	   
	   Public Function ReadProperties(bag)
			Dim x,mx
			Dim names,types
			
			mx = Cint(bag.Read("C"))
			names = Split(bag.Read("CC"),"@")
			types = Split(bag.Read("CT"),"@")
			
			If mx = 0 Then
				Exit Function
			End If
			
			For x = 0 To mx
				Controls.Add CreateWCInstance(names(x),types(x))
			Next
			
	   End Function
		
	   Public Function WriteProperties(bag)
			Dim x,mx
			Dim ctrls
			If Controls.Count = 0 Then
				Exit Function
			End If			
			mx = Controls.Count-1
			Redim ctrls(mx)
			For x = 0 To mx
				ctrls(x) = Controls.Item(x).Control.ControlID
			Next
			bag.Write "CC",Join(ctrls,"@")
			bag.Write "C",mx
	   End Function

	   Public Function ProcessPostBack()
			'Use Control.ViewState  if you want to gain access to the viewstate. 
			'It will be = Nothing EnableViewState = False.
	   End Function	    
	   
	   Public Function HandleClientEvent(e)
			HandleClientEvent = True 'To signal callee the receipt ak.
			HandleClientEvent = ExecuteEventFunctionEX(e)
			
	   End Function					
		
	   Private Function RenderMe()
			Dim x,mx
			mx = Controls.Count-1
			For x = 0 To mx
				Controls.Item(x).Render()
			Next
	   End Function
	   
	   Public Default Function Render()			
			 Dim varStart	 			 
			 If Control.IsVisible = False Then
				Exit Function
			 End If
			 
			 varStart = Now
			 
			 Call RenderMe()
			 
		 	 Page.TraceRender varStart,Now,Control.Name
		 	
	   End Function
			   	
	   	
	   Public Function AddTextBox(Name,Caption,MaxSize,Width,TextBoxType,DefaultValue)
			Dim ctrl
			Set ctrl = CreateWCInstance(Name,"ServerTextBox")
			With ctrl
			   .MaxLength = MaxSize
			   .Size   = Width
			   .Mode    = TextBoxType
			   .Text    = DefaultValue
			   .Caption = Caption
			End With
			Controls.Add ctrl
			Set AddTextBox = ctrl
	   End Function
		
	   Private Function BindToDataSource(ctrl, DataSource,DataValueField,DataTextField,DefaultValue)
			If Not DataSource Is Nothing Then
				With ctrl
					.DataSource = DataSource
					.DataTextField = DataTextField 
					.DataValueField = DataValueField
					If DefaultValue<>"" Then
					   .Value = DefaultValue
					End If
				End With
			End If
	   End Function
	   Public Function AddDropDown(Name,Caption,DataSource,DataValueField,DataTextField,DefaultValue)
			Dim ctrl
			Set ctrl = CreateWCInstance("ServerDropDown")
			BindToDataSource ctrl
			ctrl.Caption = Caption
			Controls.Add ctrl
			Set AddDropDown = ctrl
	   End Function	   

	   Public Function AddListBox(Name,Caption,Rows,DataSource,DataValueField,DataTextField,DefaultValue)			 			 
			Dim ctrl
			Set ctrl = CreateWCInstance("ServerDropDown")
			BindToDataSource ctrl
			ctrl.Caption = Caption
			ctrl.Rows = Rows
			Controls.Add ctrl
			Set AddListBox = ctrl
	   End Function
	   
	   Public Function AddCheckBoxList(Name,RepeatColumns,DataSource,DataValueField,DataTextField,DefaultValue)			 			 
			Dim ctrl
			Set ctrl = CreateWCInstance("ServerCheckBoxList")
			BindToDataSource ctrl
			ctrl.RepeatColumns= RepeatColumns
			Controls.Add ctrl
			Set AddCheckBoxList = ctrl
	   End Function
	   
	   Public Function AddRadioButtonList(Name,RepeatColumns,DataSource,DataValueField,DataTextField,DefaultValue)			 			 
			Dim ctrl
			Set ctrl = CreateWCInstance("ServerRadioButtonList")
			BindToDataSource ctrl
			ctrl.RepeatColumns= RepeatColumns
			Controls.Add ctrl
			Set AddRadioButtonList = ctrl
	   End Function

  	   Public Function AddCheckBox(Name,Caption,Checked)			 
			Dim ctrl
			Set ctrl = CreateWCInstance("ServerCheckBox")
			ctrl.Checked = Checked
			Controls.Add ctrl
			Set AddCheckBox = ctrl
	   End Function
 
	   Private Function CreateWCInstance(Name,ClassName)
			Set CreateWCInstance =  eval("New_" & ClassName & "(""" & Name & """)" )
	   End Function
	   
	   Public Function FindControl(ControlName)
			Set FindControl = Controls.Item(ControlName)
	   End Function
	   
	   Public Function GetValue(ControlName)
			ctrl = Controls.Item(ControlName)
			Select Case TypeName(ctrl)
				case "TextBox":
				   GetValue = ctrl.Text
				case "DropDown","CheckBoxList","RadioButtonList":
				   GetValue = ctrl.Value
			   case "CheckBox"
			       GetValue = ctrl.Checked
			   case else
				  Response.Write "Control : " & TypeName(ctrl) & " not supported"
				  Response.End
			End Select
	   End Function
	   
	   Public Function SetValue(ControlName,NewValue)
			ctrl = Controls.Item(ControlName)
			Select Case TypeName(ctrl)
				case "TextBox":
				   ctrl.Text  = NewValue
				case "DropDown","CheckBoxList","RadioButtonList":
				   ctrl.Value  = NewValue
			   case "CheckBox"
			       ctrl.Checked = NewValue
			   case else
				  Response.Write "Control : " & TypeName(ctrl) & " not supported"
				  Response.End
			End Select
	   End Function
				
	   Public Function RenderControl(ControlName)
		Set ctrl = Controls.Item(ControlName)
		If Not ctrl Is Nothing Then
			ctrl.Render
		End If
	   End Function	   	   
	End Class
	
%>