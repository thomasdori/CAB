<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Frame

	
	Public Function New_ServerFrame(name,width,height) 
		Set New_ServerFrame = New ServerFrame
			New_ServerFrame.Control.ControlID = name
			New_ServerFrame.Width = width
			New_ServerFrame.Height = height
	End Function

	Page.RegisterLibrary "ServerFrame"
	
	Class ServerFrame
	
		Public Control
		Public Caption
		Public CaptionStyle
		Public CaptionCssClass
		Public FrameTemplate
		Public AutoSize	
		Public Width
		Public Height
		Public OverFlow	' set overflow mode
		
		Private mHTML
		Private mContents() 	' WebControl Objects (and string, etc) to be rendered inside the layer		
				
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 
			
			Caption 		= ""
			FrameTemplate	= ""
			AutoSize		= False			
			Width			= 0
			Height			= 0
			OverFlow		= ""

			mHTML = ""
			ReDim mContents(0)
			
		End Sub
	   
	   	Public Function ReadProperties(bag)
			Caption			= bag.Read("C")
			CaptionStyle	= bag.Read("CS")
			CaptionCssClass	= bag.Read("CN")
			OverFlow		= bag.Read("O")
			mHTML			= bag.Read("T")
			Width			= bag.Read("W")
			Height			= bag.Read("H")
			AutoSize		= bag.Read("A")
			FrameTemplate	= bag.Read("F")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "C",Caption
			bag.Write "CS",CaptionStyle
			bag.Write "CN",CaptionCssClass
			bag.Write "O",OverFlow
			bag.Write "T",mHTML
			bag.Write "W",Width
			bag.Write "H",Height
			bag.Write "A",AutoSize
			bag.Write "F",FrameTemplate
		End Function
	   
		Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunctionEX(e)
		End Function

		Public Property Let InnerHTML(ByVal htm)
			mHTML = htm & ""
		End Property
		
	   	Public Function AddContent(ByRef Cn)
	   		Dim ub
	   		
	   		ub = UBound(mContents)
	   		
	   		If IsObject(cn) Then 
				' Check for WebControl
				On Error Resume Next
				If Cn.Control is Nothing Then Exit Function
				If Err.Number<>0 then Exit Function
				Set mContents(ub) = Cn	' only WebControl is to be used as an object
				'Set Cn.Control.Owner = Me	' ok to generate error if Cn is not a WebControl
			Else
				mContents(ub) = Cn		'other datatypes - string, number, etc
			End If
	   		
	   		ReDim Preserve mContents(ub+1)	   		
	   	End Function	   
	   
		Public Function OpenFrame()		
			Dim Style
			Dim FrameStyle

			Style = Control.Style
			FrameStyle =" Style='cursor:default; " + _
						IIf(Not AutoSize,"width:"& Width &"; height:"& Height &"; ","") + _
						IIf(Not Control.Visible,"display:none; ","") + _
						IIf(OverFlow<>"","overflow:" + OverFlow + "; ","") + "'"

			Response.Write vbNewLine & "<fieldset id='" & Control.ControlID  & "' "        
				Response.Write IIf(Style<>""," Style='" & Style  + "' ",FrameStyle)
				Response.Write IIf(Control.CssClass<>""," Class='" & Control.CssClass  + "' ","")
				Response.Write IIf(Not Control.Enabled," disabled='disabled' ","")
				Response.Write ">" & vbNewLine
			If Caption<>"" Then
				Response.Write "<legend "
					Response.Write IIf(CaptionStyle<>""," Style='" & CaptionStyle  + "' ","")
					Response.Write IIf(CaptionCssClass<>""," Class='" & CaptionCssClass  + "' ","")
					Response.Write ">"+ Caption +"</legend>"
			End If
		End Function
		
		Public Function CloseFrame()
			Response.Write "</fieldset>"
		End Function
		
		Public Default Function Render ()
			 Dim varStart	 
			 
			 varStart = Now
			 
			 If Control.Visible = False Then
				Exit Function
			 End If

			Call OpenFrame()
				Call RenderHTMLContent
			Call CloseFrame()
			
			Page.TraceRender varStart,Now,Control.ControlID
		End Function
		
		Private Function RenderHTMLContent()
			Dim i
			Dim fnc
			If FrameTemplate = "" Then
				If mHTML <> "" Then
					Response.Write mHTML
				End If
			Else
				On Error Resume Next
				Set fnc = GetRef(PanelTemplate)	
				On Error Goto 0
				If Not fnc Is Nothing Then
					Call fnc()
				End If
			End If
			'content array
			For i=0 to UBound(mContents)-1 ' last item is always empty
				Response.Write mContents(i) & ""
			Next			
		End Function
		
	End Class
%>
