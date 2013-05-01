<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server DropDown
	
	Public Function New_ServerDropDown(name) 
		Set New_ServerDropDown = New ServerDropDown 
			New_ServerDropDown.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerDropDown"
	
	 Class ServerDropDown
		Dim Control
		Dim Caption
		Dim mMultiple
		Dim ReadOnly
		Dim Rows
		
		Dim DataValueField
		Dim DataTextField
		Dim DataSource
		Dim Items
		Dim AutoPostBack
		Dim RaiseOnChanged
		
		Dim CaptionCssClass
		Dim CaptionStyle
		
		Private mbolWasRendered

		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Control.ImplementsProcessPostBack = True
			Control.SupportsClientSideEvent   = True
			Caption = ""			
			Rows  = 1
			mMultiple = False
				
			DataValueField  = ""
			DataTextField   = ""
			AutoPostBack    = False	
			RaiseOnChanged  = False
			mbolWasRendered = False
			ReadOnly		= False	
			
			Set DataSource =  Nothing	
			Set Items = New_ListItemsCollectionObject()						
	   End Sub
	   Private Sub Class_Terminate()
			Set Items = Nothing
			Set DataSource = Nothing
	   End Sub
	   
	   Public Property Get Multiple()
			Multiple = mMultiple
	   End Property
	   
	   Public Property Let Multiple(v)
			mMultiple = v
			Items.Mode  = IIF(mMultiple,1,0)
			If Not mMultiple Then
				Items.SetSelectedByValue Items.GetSelectedValue,True
			End If
	   End Property

	   
	   Public Property Get Text()
			Text = Items.GetSelectedText
	   End Property
	   
	   Public Property Let Text(v)
			Items.SetSelectedByText v,true	
	   End Property

	   Public Property Get Value()
			Value = Items.GetSelectedValue	
	   End Property
	   
	   Public Property Let Value(v)
			Items.SetSelectedByValue v,true
	   End Property

	   Public Function ReadProperties(bag)
	   		mbolWasRendered = CBool(bag.ReadBoolean("R"))	   			
	   		Caption         = bag.Read("C")
	   		mMultiple       = CBool(bag.ReadBoolean("M"))
	   		Rows            = CInt(bag.ReadInt("RW"))   			
	   		AutoPostBack    = CBool(bag.ReadBoolean("A"))
	   		CaptionCssClass = bag.Read("CS")
	   		CaptionStyle    = bag.Read("ST")
	   		ReadOnly        = CBool(bag.ReadBoolean("RO"))
	   		RaiseOnChanged  = CBool(bag.Read("RC"))	   		
	   		Items.SetState bag.Read("I")	
	   		Items.Mode      = IIF(mMultiple,1,0)   			   			
	   End Function

	   Public Function WriteProperties(bag)
			bag.Write "R",  False			
			bag.Write "C",  Caption
			bag.Write "M",  mMultiple
			bag.Write "RW", Rows
			bag.Write "I",  Items.GetState()
			bag.Write "CS", CaptionCssClass
	   		bag.Write "ST", CaptionStyle
	   		bag.Write "A",  AutoPostBack
	   		bag.Write "RO", ReadOnly
	   		bag.Write "RC", RaiseOnChanged
		End Function
		
	   	Public Function ProcessPostBack()
		   	Dim x,mx
			Dim OldValue
			Dim Form
			
			'If ViewState is present AND control was not rendered or was disabled then don't process.
			If Not Control.ViewState Is Nothing And Not (mbolWasRendered AND Control.Enabled)  Then
				Exit Function
			End If			

			Set Form = Request.Form(Control.ControlID)
			mx = Form.Count				
			
			If RaiseOnChanged Then
				OldValue = Items.GetSelectedValue
				If OldValue <>	Form Then
					Call Page.RegisterPostBackEventHandler(Me,"OnChanged",OldValue)
				End If		
			End If
						
			If mMultiple Then
				Items.SetAllSelected(False)
			End If

			For x = 1 To mx						  
				Call Items.SetSelectedByValue(Form.Item(x),True)
			Next				
	   	
	   	End Function


		
		Public Sub Bind(pDataSource,pDataValueField,pDataTextField,CacheAs,bolAddBlank)
			Dim bolFirstTime
			bolFirstTime = False

			If CacheAs<>"" Then
				If Application(CacheAs)<>"" Then
					bolFirstTime = True
					SetFromCache(Application(CacheAs))
					Exit Sub
				End If
			End If

			Set Me.DataSource  = pDataSource
			Me.DataValueField  = pDataValueField
			Me.DataTextField   = pDataTextField
			Me.DataBind
			
			If bolAddBlank Then
				Items.Add "","",True,0
			End If
				
			On Error Resume Next
			If CacheAs <> "" And bolFirstTime Then
				Application.Lock
				Application(CacheAs) = Items.GetState()
				Application.UnLock
				Err.Clear
			End If 
		End Sub

		Public Function SetFromCache(cache)
			Items.SetState cache
			Call ProcessPostBack()
		End Function
			   
	   	Public Function SetValueFromDataSource(value)
			If DataValueField <> "" Then
				Items.SetSelectedByValue value,True
			Else		
				Items.SetSelectedByText  value,True
			End If			
	    End Function
		
		Public Function SetValue(value) 
			Dim values
			Dim x,mx
			If mMultiple Then
				values = split(value,",")
				Items.SetAllSelected(False)
				mx = UBound(values)
				For x = 1 to mx						  
					Call Items.SetSelectedByValue(UrlDecode(values(x)),True)
				Next
			Else
				Items.SetSelectedByValue value,true
			End If
		End Function
	   
	   Public Function HandleClientEvent(e)
		 	If AutoPostBack Then
				HandleClientEvent = ExecuteEventFunctionEX(e)
		 	End If	
	   End Function			
	   
		Public Sub DataBind()
						
			Dim objRs,x,mx	
			Dim fld1,fld2
			Dim use1,use2
			
			Items.Clear()
			
			If DataSource Is Nothing Then				
				Exit Sub
			End If		

			If 	DataSource.RecordCount = 0 Then
				Exit Sub
			End If
					
			Set objRs = DataSource				
			
			objRs.MoveFirst
			
			If DataTextField<>""  Then 
				Set fld1 = objRs(DataTextField) 
				use1 = True
			Else  
				Set fld1 = Nothing
				use1 = False
			End If
			If DataValueField<>"" Then 
				Set fld2 = objRs(DataValueField) 
				use2 = True
			Else 
				Set fld2 = Nothing
				use2 = false
			End If
			x=0
			While Not objRs.EOF									
				Items.Append IIF(use1, fld1,""),  IIF(use2, fld2,x),False
				objRs.MoveNext
				x = x + 1		
			Wend

			If Control.Visible And Page.IsPostBack Then
				mx = Request.Form(Control.ControlID).Count			
				Items.SetAllSelected(False)
				For x = 1 to mx
					 Call Items.SetSelectedByValue(Request.Form(Control.ControlID).Item(x),True)
				Next
			End If
			
			Set fld1 = Nothing
			Set fld2 = Nothing
		End Sub

	   
	   Private Function RenderDropDown()
			Dim x,mx,varHTML,varInput
			Dim Selected,Value,Text
			Dim objOut 
			Dim evtName
						
			
			If AutoPostBack Then
				evtName = Page.GetEventScript("onchange",Control.ControlID,"ItemChange","this","")
			End If
			If Caption<>"" Then
				Response.Write "<span "
				Response.Write IIf(CaptionCssClass<>""," class='" & CaptionCssClass  + "' ","")
				Response.Write iif(CaptionStyle<>""," style='" & CaptionStyle + "' ","")					
				Response.Write ">" & Caption & "</span>"
			End If	
			
			If ReadOnly Then
				Response.Write  "<span id='" & Control.ControlID & "' name='" & Control.ControlID & "' " &_
					  	  IIf(Control.CssClass<>""," class='" & Control.CssClass  + "' ","") &_
				          iif(Control.Style<>""," style='" & Control.Style + "' ","") &_
				        ">" & vbNewLine												
				Response.Write Server.HTMLEncode (Items.GetSelectedText)
				Response.Write "</span>"
				Exit Function
			End If

			
			mx = Items.Count-1
			Response.Write  "<select id='" & Control.ControlID & "' name='" & Control.ControlID & "' " &_
					  	     evtName & " tabindex = " & Control.TabIndex					  	  
					  	  If Control.CssClass <> "" Then Response.Write " class='" & Control.CssClass  + "' "
				          If mMultiple  Then Response.Write " multiple "
				          If Rows > 1  Then Response.Write " size=" & Rows & " "
				          If Control.Style <> ""  Then Response.Write " style='" & Control.Style + "' "
				          If Not Control.Enabled  Then Response.Write " disabled "
				          Response.Write ">" & vbNewLine								
			For x = 0 to mx			
				Items.GetItemData x,Text,Value,Selected	
				Response.Write  "<option " & IIf(Selected," selected","") & " value='"  & Server.HTMLEncode(Value)	 & "'>" &_
							Text & "</option>" & vbNewLine						
			Next 			
			Response.Write "</select>" & vbNewLine
	   End Function

	   Public Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If

			 varStart = Now

			If Not Control.ViewState Is Nothing And Not ReadOnly Then
				Control.ViewState.Write "R",True
			End If
			 
			 If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			 End If
			 
			 Render = RenderDropDown
					
		 	 Page.TraceRender varStart,Now,Control.ControlID
		 	 
		End Function

	End Class

%>
