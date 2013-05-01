<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server CheckBox	
	
	Public Function New_ServerRadioButtonList(name) 
		Set New_ServerRadioButtonList = New ServerRadioButtonList
			New_ServerRadioButtonList.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerRadioButtonList"
	
	 Class ServerRadioButtonList
		
		Dim Control
		Dim ReadOnly
		
		Dim DataValueField
		Dim DataTextField
		Dim DataSource
		Dim Items
		Dim AutoPostBack
		Dim RaiseOnChanged
		
		Dim RepeatLayout 'Table (def)/Flow
		Dim RepeatDirection 'Vertical (def)/Horizontal)
		Dim RepeatColumns  '0 def
		
		Dim TableCss
		Dim TableStyle
		
		Dim CellCssClass
		Dim CellStyle
		
		Dim Nowrap		
		Dim BorderWidth
		Dim BorderColor
		Dim GridLines  '1 hor, 2 ver 3 both
		Dim CellSpacing 
		Dim CellPadding
		Dim ItemChanged
		Private mbolWasRendered
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Control.ImplementsProcessPostBack = True
			Control.SupportsClientSideEvent   = True			

			ReadOnly = False
			Nowrap   = False
			Set DataSource =  Nothing
			DataValueField = ""
			DataTextField  = ""
			AutoPostBack   = False	
			RaiseOnChanged = False
			ItemChanged    = False
			
			Set Items = New_ListItemsCollectionObject()			
			Items.Mode = 0 'Single Selection!!			
			TableStyle = ""	
			TableCss   = ""		
	   		CellStyle    = ""
	   		CellCssClass = ""

			BorderWidth = 0
			BorderColor = "black"
			GridLines    = 3
			CellSpacing = 0
			CellPadding = 2
			RepeatLayout   = 1
			RepeatDirection =  2
			RepeatColumns  = 1
			mbolWasRendered = False						
	   End Sub
	   Private Sub Class_Terminate()
			Set Items = Nothing
			Set DataSource = Nothing
	   End Sub

	   	Public Function ReadProperties(bag)
	   		RepeatLayOut    = Cint(bag.Read("A"))
	   		RepeatDirection = Cint(bag.Read("B"))
	   		RepeatColumns = Cint(bag.Read("C"))
	   		BorderWidth   = CInt(bag.Read("D"))
	   		BorderColor   = bag.Read("E")
	   		GridLines     = bag.Read("F")
	   		Nowrap	   	  = CBool(bag.Read("G"))
	   		TableCss      = bag.Read("H")
	   		TableStyle    = bag.Read("J")
	   		
	   		CellCssClass = bag.Read("K")
	   		CellStyle    = bag.Read("L")
	   		
			AutoPostBack     = CBool(bag.Read("M"))	   			   		
			RaiseOnChanged   = CBool(bag.Read("N"))	   		
	   		mbolWasRendered  = CBool(bag.Read("R"))
			CellSpacing = bag.Read("O")
			CellPadding = bag.Read("P")	   		
	   		Items.SetState bag.Read("I")	   				   		
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "R",False
			bag.Write "I",Items.GetState()
			
	   		bag.Write "A",RepeatLayOut 
	   		bag.Write "B",RepeatDirection 
	   		bag.Write "C",RepeatColumns 			
			bag.Write "D" ,BorderWidth
	   		bag.Write "E",BorderColor
	   		bag.Write "F",GridLines 
			bag.Write "G",Nowrap
	   		bag.Write "H",TableCss
	   		bag.Write "J",TableStyle 
	   		bag.Write "K",CellCssClass
	   		bag.Write "L",CellStyle	   		   	   		
			bag.Write "M",AutoPostBack 
			bag.Write "N",RaiseOnChanged
			bag.Write "O",CellSpacing
			bag.Write "P",CellPadding
		End Function
		
		Public Function ProcessPostBack()
			Dim CurrentValue
			Dim FormValue

			'If ViewState is present AND control was not rendered or was disabled then don't process.
			If Not Control.ViewState Is Nothing And Not (mbolWasRendered AND Control.Enabled)  Then
				Exit Function
			End If	
										
			CurrentValue = Items.GetSelectedValue				
			FormValue = Request.Form(Control.ControlID)
			If CurrentValue <> FormValue Then
				Items.SetSelectedByValue FormValue,True
				ItemChanged = True
				If RaiseOnChanged Then
					Page.RegisterPostBackEventHandler Me,"OnChanged",""
				End If
			End If

		End Function

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
			If CacheAs<>"" And bolFirstTime Then
				Application.Lock
				Application(CacheAs) = Items.GetState()
				Application.UnLock
				Err.Clear
				End If 
		End Sub

		Public Function SetFromCache(cache)
				Items.SetState cache
	   			'Only if it was rendered in the previous request (data exists in the request)
				If Page.IsPostBack Then
					Page.TraceImportantCall Me.Control, "Setting Items from Cache"
					Items.SetSelectedByValue Request.Form(Control.ControlID),True
				End If
		End Function
	   
	   
	   
	   Public Function HandleClientEvent(e)
		 	If AutoPostBack Then
				HandleClientEvent = ExecuteEventFunction(e.EventFnc)
		 	End If	
	   End Function

	   Public Function SetValue(newvalue) 
 			Me.Value = newvalue
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
				Else  Set fld1 = Nothing
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
			
			If mbolWasRendered Then
				Items.SetSelectedByValue Request.Form(Control.ControlID),True
			End If
			
			Set fld1 = Nothing
			Set fld2 = Nothing
		End Sub

		Private Function  RenderByColumn								
			Dim x,mx,i,alt
			Dim Rows,r
			Dim Selected,Value,Text
			Dim Enabled, ControlName
			Dim CellStartTag			
			Dim sCaption
			
			ControlName = Control.ControlID
						
			If RepeatLayOut = 1 Then								
				CellStartTag = "<td valign = top "
					If CellCssClass<>"" Then CellStartTag  = CellStartTag  & " class='" & CellCssClass + "' "
					If CellStyle<>"" Then CellStartTag  = CellStartTag  & " Style='" & CellStyle + "' "
					If Nowrap  Then CellStartTag  = CellStartTag  &  " NoWrap "
				CellStartTag  = CellStartTag  & ">"
			End If
			CellStartTag = CellStartTag & "<input type='radio' id='" & ControlName & "' name='" & ControlName & "' tabindex = " & Control.TabIndex & " "
			
			If Not Control.Enabled Then Enabled = " disabled "
			sCaption = ">&nbsp;<span "
			If Control.Style<>"" Then sCaption = sCaption & " style='" & Control.Style + "' "
			If Control.CssClass<>"" Then sCaption = sCaption & " class='" & Control.CssClass + "' "
			sCaption = sCaption & ">"
									
			mx  = Items.Count
			rows = Int(mx/RepeatColumns)
			
			If mx mod RepeatColumns = 0 Then
				Rows = Rows -1
			End If

			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<table cellspacing=" & CellSpacing & " cellpadding=" & CellPadding
				If TableCss<>"" Then Response.Write " class='" & TableCss & "' "
				If TableStyle<>"" Then Response.Write " style='" & TableStyle & "' "
				If BorderWidth>0 Then Response.Write " border=" & BorderWidth & " bordercolor='" & BorderColor & "'"
				Response.Write ">" & vbNewLine				
			End If
			
			i = 0						
			For r = 0 To Rows
				If RepeatLayout = 1 Then 					
					Response.Write "<tr>"
				End If								
				For x = 1 To RepeatColumns					
					If i<mx Then				
						Items.GetItemData i,Text,Value,Selected	
						Response.Write CellStartTag						
						Response.Write IIF(Selected," checked ","") &  " value=""" + Server.HTMLEncode(Value) + """ " & Enabled 
						If AutoPostBack Then
				   		   Response.Write Page.GetEventScript("onclick",Control.ControlID,"Click",i,"")
						End If								   
						Response.Write sCaption & Text & "</span>"
						If RepeatLayOut = 1 Then 
						   Response.Write "</td>"
						End If
					    i = i + 1
					End If
				Next				
				If RepeatLayout = 1 Then 
					Response.Write "</tr>" & vbNewLine
				End If
			Next
			If RepeatLayout =1 Then				
				Response.Write  "</table>" & vbNewLine
			End If
												
		End Function
	   
		Private Function  RenderByRow				
			Dim x,mx
			Dim Cols,c
			Dim Rows,r,Row,Pos
			Dim Selected,Value,Text
			Dim Enabled, sCaption,ControlName
			Dim CellStartTag

			Dim i
			ControlName = Control.ControlID
			
			If Not Control.Enabled Then Enabled = " disabled "
			sCaption = ">&nbsp;<span "
			If Control.Style<>"" Then sCaption = sCaption & " style='" & Control.Style + "' "
			If Control.CssClass<>"" Then sCaption = sCaption & " class='" & Control.CssClass + "' "
			sCaption = sCaption & ">"

			If RepeatLayOut = 1 Then								
				CellStartTag = "<td valign = top "
					If CellCssClass<>"" Then CellStartTag  = CellStartTag  & " class='" & CellCssClass + "' "
					If CellStyle<>"" Then CellStartTag  = CellStartTag  & " Style='" & CellStyle + "' "
					If Nowrap  Then CellStartTag  = CellStartTag  &  " NoWrap "
				CellStartTag  = CellStartTag  & ">"
			End If
			CellStartTag = CellStartTag & "<input type='radio' id='" & ControlName & "' name='" & ControlName & "' tabindex = " & Control.TabIndex & " "
						
			mx = Items.Count
			Cols = RepeatColumns -1
			Rows = Int(mx/RepeatColumns) - 1
			
			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<table CellSpacing=" & CellSpacing & " CellPadding=" & CellPadding
				Response.Write " id='" &  Control.ControlID & "Tbl' "
				If TableCss<>"" Then Response.Write " class='" & TableCss & "' "
				If TableStyle<>"" Then Response.Write " style='" & TableStyle & "' "
				If BorderWidth > 0 Then Response.Write " border=" & BorderWidth & " bordercolor='" & BorderColor & "'"
				Response.Write ">" & vbNewLine				
			End If
			
			Row = 0
			i = 0
			For x = 1 To 2 
				For r = 0 To Rows
					If RepeatLayout = 1 Then 					
						Response.Write "<tr>"
					End If								
					For c = 0 To Cols
						Pos = Row + (Rows * c + r + c)
						If Pos<mx Then
							Items.GetItemData Pos,Text,Value,Selected								
							Response.Write CellStartTag 
						    Response.Write IIF(Selected," checked ","") &  " value=""" + Server.HTMLEncode(Value) + """ " & Enabled 						    
						    If AutoPostBack Then
				   				Response.Write Page.GetEventScript("onclick",Control.ControlID,"Click",i,"")
				   				'Response.Write "onclick='clasp_srbl(this," & i & ")'"
						    End If		

							Response.Write sCaption & Text & "</span>"
							If RepeatLayOut = 1 Then 
							   Response.Write "</td>"
							End If
							i=i+1
						End If						
						
					Next
					If RepeatLayout = 1 Then 
						Response.Write "</tr>" & vbNewLine
					End If
				Next
				Cols = mx mod RepeatColumns
				Rows = 0
				If Cols = 0 Then 'Last Column!
					Exit For
				Else
					Row = mx-Cols
					Cols = Cols - 1
				End If
			Next
			
			If RepeatLayout =1 Then
				Response.Write  "</table>" & vbNewLine
			End If
				
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
			 
			 RepeatColumns = CInt(RepeatColumns)
			 
			 If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			 End If
			 
			 If RepeatDirection=1 Then
				Render = RenderByColumn()
			 Else
				Render = RenderByRow()
			 End If
					
		 	 Page.TraceRender varStart,Now,Control.ControlID
		 	 
		End Function

	End Class

%>