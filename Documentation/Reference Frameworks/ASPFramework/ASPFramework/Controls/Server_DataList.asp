<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server ServerDataList
	Public Function New_ServerDataList(name) 
		Set New_ServerDataList = New ServerDataList
			New_ServerDataList.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerDataList"
	
	Class ServerItemTemplate
			Public  Style
			Public  CssClass
			Public  FunctionName
			Public  Mode
			Private mRenderStart
			Private mRenderEnd
			Private fncFunctionName
			Public  ExtraAttributes		
			
			Private Sub Class_Initialize()
				Set fncFunctionName = Nothing
				mRenderStart = ""
				Mode = 1
				ExtraAttributes = ""
			End Sub
			
			Public Function Execute()
				If fncFunctionName Is Nothing Then
					Set fncFunctionName = GetFunctionReference(FunctionName)
				End If
				
				If mRenderStart = "" Then
					If Mode=1 Then						
						mRenderStart = "<TD" & IIF(CssClass<>""," Class='" & CssClass & "' "," ") & IIF(Style<>""," Style='" & Style & "' "," ") & ExtraAttributes &  " >"
						mRenderEnd   = "</TD>"
					Else
						If CssClass <> "" Or Style <>"" Then
							mRenderStart = "<SPAN " & IIF(CssClass<>""," Class='" & CssClass & "' "," ") & IIF(Style<>""," Style='" & Style & "' "," ") &  ExtraAttributes & " >"
							mRenderEnd   = "</SPAN>"
						End If
					End If					
				End If								
				Response.Write mRenderStart
				If Not fncFunctionName Is Nothing Then
					fncFunctionName()
				End If
				Response.Write mRenderEnd
			End Function

			Public Function ExecuteBind(ds)
				If fncFunctionName Is Nothing Then
					Set fncFunctionName = GetFunctionReference(FunctionName)
				End If
				
				If mRenderStart = "" Then
					If Mode=1 Then						
						mRenderStart = "<TD" & IIF(CssClass<>""," Class='" & CssClass & "' "," ") & IIF(Style<>""," Style='" & Style & "' "," ") & ExtraAttributes & " >"
						mRenderEnd   = "</TD>"
					Else
						If CssClass <> "" Or Style <>"" Then
							mRenderStart = "<SPAN " & IIF(CssClass<>""," Class='" & CssClass & "' "," ") & IIF(Style<>""," Style='" & Style & "' "," ") & ExtraAttributes & " >"
							mRenderEnd   = "</SPAN>"
						End If
					End If					
				End If			
					
				Response.Write mRenderStart
				If Not fncFunctionName Is Nothing Then
					fncFunctionName(ds)
				End If
				Response.Write mRenderEnd
			End Function
	 End Class
	
	 Class ServerDataList
		Dim Control
		Dim DataSource
		
		Dim HeaderTemplate
		Dim FooterTemplate		
		
		Dim ItemTemplate
		Dim AlternatingItemTemplate
		Dim SelectedItemTemplate
		Dim SeparatorTemplate		
		Dim EditItemTemplate	
			
		Dim SelectedItemIndex
		Dim EditItemIndex
		
		Dim RepeatLayout 'Table (def)/Flow
		Dim RepeatDirection 'Vertical (def)/Horizontal)
		Dim RepeatColumns  '0 def
		

		Private fncHeader
		Private fncFooter
		Private fncItemTemplate
		Private fncAlternatingItemTemplate
		Private fncSelectedItemTemplate
		Private fncEditItemTemplate
		Private fncSeparatorTemplate

		Public BorderWidth
		Public BorderColor
		Public GridLines  '1 hor, 2 ver 3 both
		Public CellSpacing 
		Public CellPadding
		
				
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
									
			Set HeaderTemplate			 = New ServerItemTemplate
			Set FooterTemplate			 = New ServerItemTemplate
			Set ItemTemplate             = New ServerItemTemplate
			Set AlternatingItemTemplate  = New ServerItemTemplate
			Set SelectedItemTemplate     = New ServerItemTemplate
			Set EditItemTemplate         = New ServerItemTemplate
			Set SeparatorTemplate        = New ServerItemTemplate
						
			BorderWidth = 0
			BorderColor = "black"
			GridLines    = 3
			CellSpacing = 0
			CellPadding = 2
			
			SelectedItemIndex = -1	
			EditItemIndex = -1
			
			RepeatLayout   = 1
			RepeatDirection =  1
			RepeatColumns  = 1
			
		End Sub

		Private Sub Class_Terminate()
			Set DataSource = Nothing
		End Sub

	   
	   	Public Function ReadProperties(bag)			
	   		RepeatLayOut = Cint(bag.Read("RL"))
	   		RepeatDirection = Cint(bag.Read("RD"))
	   		RepeatColumns = Cint(bag.Read("RC"))
			SelectedItemIndex = Cint(bag.Read("SI"))
			EditItemIndex = Cint(bag.Read("EI"))
			
	   		HeaderTemplate.FunctionName = bag.Read("HT")
	   		FooterTemplate.FunctionName  = bag.Read("FT")
	   		ItemTemplate.FunctionName = bag.Read("IT")
	   		AlternatingItemTemplate.FunctionName = bag.Read("AIT")
	   		EditItemTemplate.FunctionName = bag.Read("EIT")
	   		SelectedItemTemplate.FunctionName = bag.Read("SIT")
			SeparatorTemplate.FunctionName = bag.Read("ST")
			
			HeaderTemplate.Style = bag.Read("HTS")
			FooterTemplate.Style = bag.Read("FTS")
	   		ItemTemplate.Style = bag.Read("ITS")
	   		AlternatingItemTemplate.Style = bag.Read("AITS")
	   		EditItemTemplate.Style = bag.Read("EITS")
	   		SelectedItemTemplate.Style = bag.Read("SITS")
			
			BorderWidth = CInt(bag.Read("W"))
	   		BorderColor = bag.Read("BC")
	   		GridLines   = bag.Read("GL")
			
		End Function
		
		Public Function WriteProperties(bag)
	   		bag.Write   "RL",RepeatLayOut 
	   		bag.Write   "RD",RepeatDirection 
	   		bag.Write   "RC",RepeatColumns 
			bag.Write   "SI",SelectedItemIndex
			bag.Write   "EI",EditItemIndex
			
	   		bag.Write   "HT",HeaderTemplate.FunctionName
	   		bag.Write   "FT",FooterTemplate.FunctionName
	   		bag.Write   "IT",ItemTemplate.FunctionName
	   		bag.Write   "AIT",AlternatingItemTemplate.FunctionName
	   		bag.Write   "EIT",EditItemTemplate.FunctionName
	   		bag.Write   "SIT",SelectedItemTemplate.FunctionName
			bag.Write   "ST",SeparatorTemplate.FunctionName

	   		bag.Write   "HTS",HeaderTemplate.Style
	   		bag.Write   "FTS",FooterTemplate.Style
	   		bag.Write   "ITS",ItemTemplate.Style
	   		bag.Write   "AITS",AlternatingItemTemplate.Style
	   		bag.Write   "EITS",EditItemTemplate.Style
	   		bag.Write   "SITS",SelectedItemTemplate.Style
	   		
			bag.Write "W",BorderWidth
	   		bag.Write "BC",BorderColor
	   		bag.Write "GL",GridLines 


		End Function
	   
	    Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunctionEX(e)
	    End Function			

		Private Function  RenderByColumn								
			Dim x,mx,i,alt
			Dim Rows,r
			i=1			
			mx = DataSource.RecordCount							
			rows = Int(mx/RepeatColumns)
			
			If mx mod RepeatColumns = 0 Then
				Rows = Rows -1
			End If
			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<table cellspacing=" & CellSpacing & " cellpadding=" & CellPadding
				Response.Write " id='" &  Control.ControlID & "' "
				Response.Write IIF(Control.CssClass<>""," Class='" & Control.CssClass & "' ","")
				Response.Write IIF(Control.Style<>""," Style='" & Control.Style & "' ","")
				If BorderWidth > 0 Then
					Response.Write " Border=" & BorderWidth & " BorderColor='" & BorderColor & "'"
				End If
				Response.Write ">" & vbNewLine				
				If HeaderTemplate.FunctionName<>"" Then
					Response.Write "<TD Colspan=" & RepeatColumns
					HeaderTemplate.Execute()
					Response.Write "</TD>"
				End If
			Else
				HeaderTemplate.Execute
			End If			
			DataSource.MoveFirst			
			'Process
			For r = 0 To Rows
				If RepeatLayout = 1 Then 					
					Response.Write "<TR>"
				End If								
				For x = 1 To RepeatColumns					
					If Not DataSource.EOF Then				
						
						If DataSource.AbsolutePosition = SelectedItemIndex Then	
							SelectedItemTemplate.ExecuteBind(DataSource)
						ElseIf DataSource.AbsolutePosition = EditItemIndex Then
							EditItemTemplate.ExecuteBind(DataSource)
						Else
							If alt = 0 Then
	    						ItemTemplate.ExecuteBind(DataSource)	
	    					Else
	    						AlternatingItemTemplate.ExecuteBind(DataSource)	
	    					End If							
	    				End If	  
	    				DataSource.MoveNext  			    						    										
					End If
					alt = 1 - alt				
				Next				
				If RepeatLayout = 1 Then 
					Response.Write "</TR>"
				End If
			Next
			If FooterTemplate.FunctionName<>"" Then
				If RepeatLayout =1 Then				
						Response.Write "<TR><TD Colspan=" & RepeatColumns
						FooterTemplate.Execute()
						Response.Write "</TD></TR>"
						Response.Write  "</TABLE>" & vbNewLine
				Else
					FooterTemplate.Execute()
				End If
			End If
									
		End Function

		Private Function  RenderByRow
			Dim x,mx,i,alt
			Dim Cols,c
			Dim Rows,r,Row,Pos
			
			mx = DataSource.RecordCount							
			Cols = RepeatColumns -1
			Rows = Int(mx/RepeatColumns) - 1

			
			
			If RepeatLayout =1 Then
				Response.Write  vbNewLine &  "<TABLE CellSpacing=" & CellSpacing & " CellPadding=" & CellPadding
				Response.Write " id='" &  Control.ControlID & "' "
				Response.Write IIF(Control.CssClass<>""," Class='" & Control.CssClass & "' ","")
				Response.Write IIF(Control.Style<>""," Style='" & Control.Style & "' ","")
				If BorderWidth > 0 Then
					Response.Write " Border=" & BorderWidth & " BorderColor='" & BorderColor & "'"
				End If
				Response.Write ">" & vbNewLine				
				If HeaderTemplate.FunctionName<>"" Then
					Response.Write "<TD Colspan=" & RepeatColumns 
					HeaderTemplate.Execute()
					Response.Write "</TD>"
				End If
			Else
				HeaderTemplate.Execute
			End If
			
			DataSource.MoveFirst			
			Row = 0
			For x = 1 To 2 
				For r = 0 To Rows
					If RepeatLayout = 1 Then 					
						Response.Write "<TR>"
					End If								
					For c = 0 To Cols
						Pos = Row + (Rows * c + r + c)+1
						If Pos<=mx Then
							DataSource.AbsolutePosition = Pos
						End If
						
						If Rows<mx Then												
							If DataSource.AbsolutePosition = SelectedItemIndex Then							
								SelectedItemTemplate.ExecuteBind(DataSource)
							ElseIf DataSource.AbsolutePosition = EditItemIndex Then
								EditItemTemplate.ExecuteBind(DataSource)
							Else
								If alt = 0 Then
	    							ItemTemplate.ExecuteBind(DataSource)	
	    						Else
	    							AlternatingItemTemplate.ExecuteBind(DataSource)	
	    						End If
	    					End If	    				    				
						End If						
						alt = 1 - alt				
					Next
					
					If RepeatLayout = 1 Then 
						Response.Write "</TR>"
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
			If FooterTemplate.FunctionName<>"" Then
				If RepeatLayout =1  Then
						Response.Write "<TR><TD Colspan=" & RepeatColumns
						FooterTemplate.Execute()
						Response.Write "</TD></TR>"
						Response.Write  "</TABLE>" & vbNewLine
				Else
					FooterTemplate.Execute
				End If
			End If			
		End Function
	   	    
	    Public Default Function Render()	
				Dim varStart	 
			
				If Control.Visible = False Then	
					Exit Function
				End If
					
				varStart = Now
			 
				RepeatColumns = CInt(RepeatColumns)
					    		
	    		If AlternatingItemTemplate.FunctionName = "" Then
	    			AlternatingItemTemplate.FunctionName = ItemTemplate.FunctionName
	    		End If
	    		
	    		If EditItemTemplate.FunctionName = "" Then
	    			EditItemTemplate.FunctionName = ItemTemplate.FunctionName
	    		End If

	    		If SelectedItemTemplate.FunctionName = "" Then
	    			SelectedItemTemplate.FunctionName = ItemTemplate.FunctionName
	    		End If

				HeaderTemplate.Mode=RepeatLayout
				FooterTemplate.Mode=RepeatLayout
				ItemTemplate.Mode=RepeatLayout
				AlternatingItemTemplate.Mode=RepeatLayout
			 	EditItemTemplate.Mode=RepeatLayout
			 	SelectedItemTemplate.Mode=RepeatLayout
			 				 	
			 	If Not DataSource Is Nothing Then
			 		DataSource.MoveFirst
			 	End If
			 	
			 	If RepeatDirection = 1 Then
					RenderByColumn()
				Else
					RenderByRow()
				End If

				Page.TraceRender varStart,Now,Control.ControlID
				
	    End Function

	End Class

%>