<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server CheckBox	
	
	Public Function New_ServerCheckBoxList(name) 
		Set New_ServerCheckBoxList = New ServerCheckBoxList
			New_ServerCheckBoxList.Control.Name = name
	End Function
	
	 Page.RegisterLibrary "ServerCheckBoxList"
	
	 Class ServerCheckBoxList
		
		Dim Control
		Dim ReadOnly
		
		Dim DataValueField
		Dim DataTextField
		Dim DataSource
		Dim Items
		Dim AutoPostBack

		Dim RepeatLayout 'Table (def)/Flow
		Dim RepeatDirection 'Vertical (def)/Horizontal)
		Dim RepeatColumns  '0 def
		
		Dim TableCss
		Dim TableStyle
		Dim BorderWidth
		Dim BorderColor
		Dim GridLines  '1 hor, 2 ver 3 both
		Dim CellSpacing 
		Dim CellPadding

		
		'Render State
		Private mbolWasRendered
					
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			
			Control.ImplementsProcessPostBack = True
			Control.SupportsClientSideEvent   = True			
			
			ReadOnly = False
		
			Set DataSource =  Nothing
			DataValueField = ""
			DataTextField = ""
			AutoPostBack = False	
			
			Set Items = New_ListItemsCollectionObject()			
			Items.Mode      = 2
			TableStyle      = ""	
			TableCss        = ""		
			BorderWidth     = 0
			BorderColor     = "black"
			GridLines       = 3
			CellSpacing     = 0
			CellPadding     = 2
			RepeatLayout    = 1
			RepeatDirection =  2
			RepeatColumns   = 1								
			mbolWasRendered = False
	   End Sub

	   Private Sub Class_Terminate()
			Set Items = Nothing
			Set DataSource = Nothing
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			mbolWasRendered = CBool(bag.Read("R"))
	   		RepeatLayOut = Cint(bag.Read("A"))
	   		RepeatDirection = Cint(bag.Read("B"))
	   		RepeatColumns = Cint(bag.Read("C"))
	   		BorderWidth = CInt(bag.Read("D"))
	   		BorderColor = bag.Read("E")
	   		GridLines   = bag.Read("F")
	   		TableStyle   = bag.Read("G")
	   		TableCss     = bag.Read("H")
	   		AutoPostBack     = CBool(bag.Read("L"))	   			   		
	   		Items.SetState bag.Read("I")	   				   		
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write  "R",False
			bag.Write "I",Items.GetState()
	   		bag.Write "A",RepeatLayOut 
	   		bag.Write "B",RepeatDirection 
	   		bag.Write "C",RepeatColumns 			
			bag.Write "D" ,BorderWidth
	   		bag.Write "E",BorderColor
	   		bag.Write "F",GridLines 

	   		bag.Write "G",TableStyle
	   		bag.Write "H",TableCss
	   		bag.Write "L",AutoPostBack  
		End Function
		
		Public Function ProcessPostBack()
			Dim x,mx
			Dim frmElement

			'If ViewState is present AND control was not rendered or was disabled then don't process.
			If Not Control.ViewState Is Nothing And Not (mbolWasRendered AND Control.Enabled)  Then
				Exit Function
			End If			
		
			Items.SetAllSelected(False)
			Set frmElement = Request.Form(Control.ControlID)
			mx = frmElement.Count
			For x = 1 to mx						  
				Call Items.SetSelectedByValue(frmElement.Item(x),True)
			Next				
			
		End Function
	   
	    Public Function HandleClientEvent(e)
		 	If AutoPostBack Then
				HandleClientEvent = ExecuteEventFunctionEX(e)				
		 	End If	
	   End Function
	   
		Public Function SetValue(value) 
			Dim values
			Dim x,mx
			values = split(value,",")
			Items.SetAllSelected(False)
			mx = UBound(values)
			For x = 0 to mx						  
				Call Items.SetSelectedByValue(UrlDecode(values(x)),True)
			Next				
		End Function

		Public Function SetFromCache(cache)
	   			Dim x,mx

				Items.SetState cache
				'If mbolWasRendered Then
					Page.TraceImportantCall Me.Control, "Setting Items from Cache"
					mx = Request.Form(Control.ControlID).Count			
					Items.SetAllSelected(False)
					For x = 1 to mx						  
						 Call Items.SetSelectedByValue(Request.Form(Control.ControlID).Item(x),True)
						 Exit For 'Only once!
					Next
				'End If
				
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
				mx = Request.Form(Control.ControlID).Count			
				Items.SetAllSelected(False)										
				For x = 1 to mx						  
					 Call Items.SetSelectedByValue(Request.Form(Control.ControlID).Item(x),True)
				Next
			End If
			Set fld1 = Nothing
			Set fld2 = Nothing
		End Sub

		Private Function  RenderByColumn								
			Dim x,mx,i,alt
			Dim Rows,r
			Dim Selected,Value,Text
			Dim Enabled, ControlName,Style,Css
			Dim sCaption
			Dim ChkStartTag
					
			Enabled = IIf(Not Control.Enabled," disabled ","") 
			Style = IIf(Control.Style<>""," style='" & Control.Style + "' ","")
			Css = IIf(Control.CssClass<>""," class='" & Control.CssClass + "' ","")
			ControlName = Control.ControlID
			sCaption = ">&nbsp;<span " & Style & Css & ">"
									
			mx  = Items.Count
			rows = Int(mx/RepeatColumns)
			
			If mx mod RepeatColumns = 0 Then
				Rows = Rows -1
			End If

			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<table cellspacing=" & CellSpacing & " cellpadding=" & CellPadding
				If TableCss<>"" Then Response.Write " class='" & TableCss & "' "
				If TableStyle<>"" Then Response.Write " style='" & TableStyle & "' "
				If BorderWidth > 0 Then
					Response.Write " border=" & BorderWidth & " bordercolor='" & BorderColor & "'"
				End If
				Response.Write ">" & vbNewLine				
			End If

			ChkStartTag = "<input type='checkbox' id='" & ControlName & "' name='" & ControlName & "' "
			
			i = 0						
			For r = 0 To Rows
				If RepeatLayout = 1 Then 					
					Response.Write "<tr>"
				End If								
				For x = 1 To RepeatColumns					
					If i<mx Then				
						Items.GetItemData i,Text,Value,Selected	
						If RepeatLayOut = 1 Then 
							Response.Write "<td>"
						End If						
						   Response.Write ChkStartTag
						   Response.Write IIf(Selected," checked "," ") & " value = """ + Server.HTMLEncode(Value) + """ " &  Enabled
						   If AutoPostBack Then
				   				Response.Write Page.GetEventScript("onclick",ControlName,"Click",i,"")	   		
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
			Dim Enabled, Style, Css,sCaption,ControlName
			Dim ChkStartTag
			
			
			Enabled = IIf(Not Control.Enabled," disabled ","") 
			Style = IIf(Control.Style<>""," style='" & Control.Style + "' ","")
			Css = IIf(Control.CssClass<>""," class='" & Control.CssClass + "' ","")
			ControlName = Control.ControlID
			sCaption = ">&nbsp;<span " & Style & Css & ">"
			
			mx = Items.Count
			Cols = RepeatColumns -1
			Rows = Int(mx/RepeatColumns) - 1

			ChkStartTag = "<input type='checkbox' id='" & ControlName & "' name='" & ControlName & "' "

			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<table cellspacing=" & CellSpacing & " cellpadding=" & CellPadding
				Response.Write " id='" &  Control.ControlID & "tbl' "
				If TableCss<>"" Then Response.Write " class='" & TableCss & "' "
				If TableStyle<>"" Then Response.Write " style='" & TableStyle & "' "
				If BorderWidth > 0 Then	Response.Write " border=" & BorderWidth & " borderColor='" & BorderColor & "'"
				Response.Write ">" & vbNewLine				
			End If
			Row = 0
			For x = 1 To 2 
				For r = 0 To Rows
					If RepeatLayout = 1 Then 					
						Response.Write "<tr>"
					End If								
					For c = 0 To Cols
						Pos = Row + (Rows * c + r + c)
						If Pos<mx Then
							Items.GetItemData Pos,Text,Value,Selected	
							
							If RepeatLayOut = 1 Then 
								Response.Write "<td>"
							End If							
							   Response.Write ChkStartTag
							   Response.Write IIf(Selected," checked "," ") & " value = """ + Server.HTMLEncode(Value) + """ " & Enabled
							   If AutoPostBack Then
				   					Response.Write Page.GetEventScript("onclick",ControlName,"Click",Pos,"")
							   End If
							   Response.Write sCaption & Text & "</span>"
							   If RepeatLayOut = 1 Then 
									Response.Write "</td>"
							   End If
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