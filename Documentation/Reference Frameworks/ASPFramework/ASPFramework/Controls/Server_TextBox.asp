<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server TextBox
'::Original By Christian Calderon
'::Enhancements By Raymond Irving (Modes 5, 6 and 7)

	Dim CSS_Name_ServerTextBox
	CSS_Name_ServerTextBox = ""
	
	'Helper function.
	Public Function New_ServerTextBox(name) 
		Set New_ServerTextBox = New ServerTextBox
			New_ServerTextBox.Control.Name = name
	End Function

	Public Function New_ServerTextBoxEX(name,maxsize,width)	
		Set New_ServerTextBoxEX = New ServerTextBox
		New_ServerTextBoxEX.Control.Name = name		
		New_ServerTextBoxEX.Size=width
		New_ServerTextBoxEX.MaxLength = maxsize 
	End Function

	Public Function New_ServerDateTextBox(name) 
		Set New_ServerDateTextBox = New ServerTextBox
			New_ServerDateTextBox.Control.Name = name
			New_ServerDateTextBox.Mode = 8
		Page.RegisterClientImage "ServerCalendar",SCRIPT_LIBRARY_PATH +"calendar/images/cal.gif"			
	End Function

	Public Function New_ServerTextArea(name,cols,rows) 
		Set New_ServerTextArea = New ServerTextBox
			New_ServerTextArea.Control.Name = name
			New_ServerTextArea.Mode=3
			New_ServerTextArea.Cols = cols
			New_ServerTextArea.Rows = rows
	End Function
	
	'Flags for load once JavaScripts
	Public AlreadyLoadFNScript : AlreadyLoadFNScript = False
	Public AlreadyLoadCALScript: AlreadyLoadCALScript = False
	
	Page.RegisterLibrary "ServerTextBox"
	
	 Class ServerTextBox
		
		Dim Control
		Dim Text
		Dim Caption
		Dim Mode '1 Text, 2 Password, 3 TextArea, 4 Label, 5 Upper, 6 Lower, 7 Numeric, 8 Date
		Dim FormatString 'only to be used in mode 7
		Dim Size
		Dim ReadOnly
		Dim MaxLength
		Dim Cols
		Dim Rows
		Dim TextChanged
		Dim RaiseOnChanged
		Dim AutoPostBack
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 						

			Control.ImplementsProcessPostBack = True
			Control.SupportsClientSideEvent   = True
				
			Text    = ""
			Caption = ""
			Mode    = 1
			ReadOnly = False
			Size=0
			Control.CssClass = CSS_Name_ServerTextBox
			MaxLength = 0
			Cols = 40
			Rows = 5			
			TextChanged = False			
			RaiseOnChanged = False
			FormatString = ""
			AutoPostBack = False
	   End Sub

		Public Function WriteProperties(bag)					
			bag.Write "C",Caption',""
			bag.Write "FS",FormatString',""
			bag.Write "T",Text',""			
			bag.Write "P",Mode & ";" & Size & ";" & MaxLength & ";" & CInt(ReadOnly) & ";" & Rows & ";" & Cols & ";" & CInt(RaiseOnChanged) & ";" & CInt(AutoPostBack)
		End Function

	   	Public Function ReadProperties(bag)			
			Dim varData
			Caption   = bag.Read("C")
			FormatString = bag.Read("FS")			
			varData = Split(bag.Read("P"),";")
			Mode	  = CInt(varData(0))			
			Size	  = CInt(varData(1))
			MaxLength = CInt(varData(2))
			ReadOnly  = CBool(varData(3))
			Rows	  = CInt(varData(4))
			Cols      = CInt(varData(5))
			RaiseOnChanged = CBool(varData(6))
			AutoPostBack = CBool(varData(7))
		End Function
		
		Public Function ProcessPostBack()
			If  Request.Form.Key(Control.ControlID) <> "" Then
				Text  = Request.Form(Control.ControlID)
				If Not Control.ViewState Is Nothing Then
					TextChanged = (Control.ViewState.Read("T")<>Text)
				End If
			Else
				If Not Control.ViewState Is Nothing Then
					Text  = Control.ViewState.Read("T")
				End If
			End If
		
			If RaiseOnChanged  Then				
				If TextChanged Then
					Call Page.RegisterPostBackEventHandler(Me,"OnChanged","")
				End If
			End If
		End Function
			   	   
	   Public Function HandleClientEvent(e)
		 	HandleClientEvent = False	
	   End Function					
	   
	   Public Function SetValueFromDataSource(value)
			Text = value
	   End Function
	   
	   Public Function SetValue(value) 			
			Text = value
	   End Function
			   
	   Private Function RenderLabel()
	   	   		Response.Write  "<span id='" & Control.ControlID & "' name='" &_
			       Control.ControlID & "' " &_
			       IIf(Control.CssClass<>""," class='" & Control.CssClass  + "' ","") &_
			       IIf(Control.Style<>""," style='" & Control.Style + "' ","")  & ">"
			   Response.Write Server.HTMLEncode(Text)
			   Response.Write "</span>"			    
	   End Function
	   
	   Private Function RenderDateTextBox()
			
			If Not AlreadyLoadCALScript Then 				
				AlreadyLoadCALScript = True ' calendar script
				Response.Write "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "calendar/calendar.js'></script>"
			End If
			
			Response.Write  "<INPUT  id='" & Control.ControlID & "' name='" &_
			       Control.ControlID & "' Value=""" & Server.HTMLEncode(Text) & """  TYPE=""TEXT"" MAXLENGTH = ""10"" "  &_
			       " TabIndex = " & Control.TabIndex
					If Control.CssClass<>""		Then Response.Write " class='" & Control.CssClass  + "' "
					If Control.Style<>""		Then Response.Write " style='" & Control.Style + "' "
					If Control.Attributes <> "" Then Response.Write Control.Attributes   & " "
					If Size > 0				   Then Response.Write " size=" & Size & " "
					If ReadOnly				   Then Response.Write " readonly "
					If Not Control.Enabled	   Then Response.Write " disabled  "			       			       
					Response.Write ">"				
			If Not ReadOnly Then 
				Response.Write "&nbsp;<a href='javascript:showCalendar(document." & Control.ControlID & ");'><img alt='Click here to select date' src='" + SCRIPT_LIBRARY_PATH + "calendar/images/cal.gif' border=0 width='16' height='16' align = 'middle'></a>"
			End If
			Page.RegisterClientScript "Calendar" & Control.ControlID ,  "<script language='JavaScript'> document." &  Control.ControlID &  " = new calendar(clasp.form.getField('" & Control.ControlID & "'));</script>"
	   End Function	   
	   
	   Private Function RenderTextBox()
		   	Dim jsEvent
			If Mode <> 1 Then
				Select Case Mode 
				Case 5 : 'Upper
					Text = UCase(Text)
					jsEvent = " onchange = ""this.value=(this.value+'').toUpperCase()"" "
				Case 6 :'Lower
					Text = LCase(Text)
					jsEvent = " onchange = ""this.value=(this.value+'').toLowerCase()"" "
				Case 7 : 'Numeric  to-do: add regex to support decimal "."
					If IsNumeric(Text) Then 
						Text = UCase(Text)  
					Else 
						Text =""
					End If			
					jsEvent = " onchange = ""this.value=CNumber(this.value,'"& FormatString &"')"" "
					If Not AlreadyLoadFNScript Then 				
						AlreadyLoadFNScript = True	' format Number script
						Response.Write "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "formatNumber.js'></script>"
					End If
				End Select
			End If
			
	   		Response.Write  "<input id='" & Control.ControlID & "' name='" &_
			       Control.ControlID & "' value=""" & Server.HTMLEncode(Text) & """ "  &_
			       " tabindex = " & Control.TabIndex
			       If Control.CssClass	 <> "" Then Response.Write " class='" & Control.CssClass  & "' "
			       If Control.Style		 <> "" Then Response.Write " style='" & Control.Style	  & "' "
			       If Control.Attributes <> "" Then Response.Write Control.Attributes   & " "
			       If Size > 0				   Then Response.Write " size=" & Size & " "
			       If MaxLength > 0			   Then Response.Write " maxlength=" & MaxLength & " "
			       If ReadOnly				   Then Response.Write " ReadOnly "
			       
			       If Mode = 2 Then
						Response.Write " type='password' "
			       Else
						Response.Write " type='text' "  & jsEvent
				   End If
				   				   			   
			       If Not Control.Enabled	   Then Response.Write " Disabled  "			       			     
				   Response.Write ">"							     			     
	   End Function

	   Private Function RenderTextArea()									   		
	   		Response.Write  "<textarea id='" & Control.ControlID & "' name='" & Control.ControlID & "' " &_
			       " TabIndex = " & Control.TabIndex & " rows = " & Rows  & " cols = " & Cols  
			       If Control.CssClass<>""		Then Response.Write " class='" & Control.CssClass  + "' "
			       If Control.Attributes<>""	Then Response.Write Control.Attributes  + " "
			       If Control.Style<>""			Then Response.Write " style='" & Control.Style + "' "
			       If MaxLength > 0				Then Response.Write " maxlength=" & MaxLength & " "
			       If ReadOnly					Then Response.Write " readonly "
			       If Not Control.Enabled	   Then Response.Write " disabled  "			       			       
				   Response.Write ">"							     			     
			Response.Write Server.HTMLEncode(Text)
			Response.Write "</textarea>"
	   End Function

	
	   Public  Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.IsVisible = False Then
				Exit Function
			 End If

			 varStart = Now
			
			 If Control.TabIndex = 0 Then  'If Not assigned, then autoassign
				Control.TabIndex = Page.GetNextTabIndex()
			 End If
			 
			 If Caption<>"" Then
			 	If Left(Caption,1) = "^" Then Caption = Mid(Caption,2) + "<br>" Else Caption = Caption & "&nbsp;"
				Response.Write Caption 
			 End If

			 Select Case Mode
				Case 3:
					Render = RenderTextArea
				Case 4:
					Render = RenderLabel
				Case 8:
					Render = RenderDateTextBox
				Case Else
					Render = RenderTextBox				
			 End Select		 			
			 If AutoPostBack Then
				Page.RegisterEventListener Me,"onchange","clasp.form.doPostBack(""OnChanged"",""" & Control.ControlID & """);"
			 End If
			 		
		 	 Page.TraceRender varStart,Now,Control.ControlID
		 	 
		End Function

	End Class

%>
