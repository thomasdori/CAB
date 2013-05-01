<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Layer

	Public Function New_ServerLayer(name,html,left,top,width,height,bgcolor) 
		Set New_ServerLayer = New ServerLayer
			New_ServerLayer.Control.Name = name
			New_ServerLayer.InnerHTML = html
			New_ServerLayer.Left = left
			New_ServerLayer.Top = top
			New_ServerLayer.Width = width
			New_ServerLayer.Height = height
			New_ServerLayer.BGColor = bgcolor
	End Function

	'Flag to only allow one richtext script to be loaded
	Public AlreadyLoadDragDropScript : AlreadyLoadDragDropScript = False	
	
	Page.RegisterLibrary "ServerLayer"
	
	Class ServerLayer
		Public Control
		Public Top
		Public Left
		Public Width
		Public Height
		Public BGColor
		Public DragEnabled
		Public LayerTemplate
		Public ZIndex
		Public AutoPostBack
		Public AutoSize
		Public PreventBubble	' prevent mouse events from being passed to parent layer/document
		Public Position
		Public RaiseOnMove
		
		Public ParentLayer		' Parent Layer
		Public ExtraAttributes	' extra attributes for <div>
		
		Private mHtml
		Private mContents() 	' WebControl Objects (and string, etc) to be rendered inside the layer			
		
		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me
				Control.ImplementsProcessPostBack = True

			'Initialize internal variables
			Top				= 0
			Left			= 0
			Width			= 100
			Height			= 100
			ZIndex			= 0
			BGColor			= ""
			DragEnabled		= False
			LayerTemplate	= ""
			AutoPostBack 	= False
			AutoSize		= False
			PreventBubble	= False
			Position		="absolute"
			RaiseOnMove		= False

			ExtraAttributes	= ""
			Set ParentLayer	= Nothing
			
			mHTML 			= ""
			ReDim mContents(0)
			
		End Sub
	   		
		Public Function ReadProperties(bag)
			Top				= bag.Read("T")
			Left			= bag.Read("L")
			Width			= bag.Read("W")
			Height			= bag.Read("H")
			ZIndex			= bag.Read("Z")
			BGColor			= bag.Read("C")
			DragEnabled		= bag.Read("D")
			mHTML 			= bag.Read("M")
			LayerTemplate	= bag.Read("R")
			AutoPostBack	= bag.Read("P")
			AutoSize		= bag.Read("A")
			PreventBubble	= bag.Read("B")
			Position		= bag.Read("O")
			RaiseOnMove		= CBool(bag.Read("V"))
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",Top
			bag.Write "L",Left
			bag.Write "W",Width
			bag.Write "H",Height
			bag.Write "Z",ZIndex
			bag.Write "C",BGColor
			bag.Write "D",DragEnabled
			bag.Write "M",mHTML
			bag.Write "R",LayerTemplate
			bag.Write "P",AutoPostBack
			bag.Write "A",AutoSize
			bag.Write "B",PreventBubble
			bag.Write "O",Position			
			bag.Write "V",RaiseOnMove
		End Function

		Public Function ProcessPostBack()
			Dim vl
			Dim queue
			Dim PosChanged
			vl = Request.Form(Control.ControlID+"_Value")
			If vl <> "" Then
				queue = Split(vl,";")
				PosChanged = (vl <> (Left &";"& Top &";"& ZIndex))
				If UBound(queue)>1 Then
					Left = CInt(queue(0))
					Top = CInt(queue(1))
					ZIndex = CInt(queue(2))
				End If
			End If

			If RaiseOnMove Then				
				If PosChanged Then
					Call Page.RegisterPostBackEventHandler(Me,"OnMove","")
				End If
			End If			
		End Function
		
		Public Function HandleClientEvent(e)
			If AutoPostBack Then
				HandleClientEvent = ExecuteEventFunction(e.EventFnc)
			End If
		End Function					

		Public Property Get InnerHTML()
			InnerHTML = mHTML
		End Property
		
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
				Set Cn.ParentLayer = Me	' ok to generate error if Cn is not a Layer
			Else
				mContents(ub) = Cn		'other datatypes
			End If
	   		
	   		ReDim Preserve mContents(ub+1)
	   		
	   	End Function	   

		Public Default Function Render()			
			 Dim varStart	 

			 If Control.IsVisible = False Then
				Exit Function
			 End If

			 varStart = Now

			 Call OpenLayer()
			 	RenderHTMLContent
			 Call CloseLayer()

			 Page.TraceRender varStart,Now,Control.Name

		End Function
		
		Public Function OpenLayer()
			Dim lStyle
			Dim sTag
			Dim eTag
			Dim evtName
			Dim jsEvtArgs

			If Not AlreadyLoadDragDropScript Then
				AlreadyLoadDragDropScript = True
				Response.Write "<script language=""JavaScript"" src=""" + SCRIPT_LIBRARY_PATH + "dragdrop/dragdrop.js""></script>"
			End If

			'temp - disable dragging if position not set to absolute
			If LCase(Position)<>"absolute" Then	
				DragEnabled = False
			End If
			
	   		If AutoPostBack Then
	   			evtName = Page.GetEventScript("onclick",Control.ControlID,"Click","","")
	   		End If
	   		
	   		jsEvtArgs = "{" + _
	   					"e:event," +_
	   					"elm:this," +_
	   					"ctrl:'"+Control.Name+"_Value',"+ _
	   					"pb:"+IIf(PreventBubble Or DragEnabled,"true","false") + _
	   					"}"
	   		
			If BGColor ="" Then BGColor="null"
			lStyle = Control.Style +";" + _
				"visibility: visible;" + _
				IIf(Not AutoSize,"clip: rect(0px "+CStr(Width)+"px "+CStr(Height)+"px 0px);","") + _
				"position:"+Position+";" + _
				"z-index: "+CStr(ZIndex)+";" + _
				"background-color: "+BGColor+";" + _
				"layer-background-color: "+BGColor+";"
			
			Response.Write "<input name='"+Control.Name+"_Value' type='hidden' value='"+CStr(Left)+";"+CStr(Top)+";"+CStr(ZIndex)+"'>"
			If (Browser.IsNS And Browser.Version=4) Then			
				Response.Write "<style> #"+Control.Name+" {"+lStyle+"}</style>"
				Response.Write "<layer id='"+Control.Name+"' "
					If Control.CssClass Then Response.Write "class='"+Contron.CssClass+"' "
					Response.Write "left='"+CStr(Left)+"' top='"+CStr(Top)+"' "+ _
									IIf(Not AutoSize,"width='"+CStr(Width)+"' height ='"+CStr(Height)+"' ","")
					If DragEnabled Then
						Response.Write "onload=""initNS4("+jsEvtArgs+")"" "
					End If
					Response.Write ExtraAttributes+" "+evtName+">"
			Else
				Response.Write "<div id='"+Control.Name+"' "
					If Control.CssClass<>"" Then Response.Write "class='" & Contron.CssClass+"' "
					Response.Write "style= '"+lStyle +"left:"+CStr(Left)+"px; top:"+CStr(Top)+"px; " + _
									IIf(Not AutoSize,"width:"+CStr(Width)+"px; height:"+CStr(Height)+"px; ","")+"' "
					If DragEnabled Then
						Response.Write "onmousedown=""grabElm("+jsEvtArgs+")"" "
						Response.Write "onmouseup=""dropElm("+jsEvtArgs+")"" "
						Response.Write "onmousemove=""moveElm("+jsEvtArgs+")"" "
						Response.Write "onselectstart = ""checkElm()"" "
					ElseIf PreventBubble Then
						If Not AutoPostBack Then Response.Write "onclick=""pbElm(event)"" "
						Response.Write "onmousedown=""pbElm(event)"" onmouseup=""pbElm(event)"" onmousemove=""pbElm(event)"" onselectstart=""checkElm()"" "
					End If
					Response.Write ExtraAttributes+" "+evtName+">"
			End If
		End Function			

		Public Function CloseLayer()
			If (Browser.IsNS And Browser.Version=4) Then
				Response.Write "</layer>"
			Else
				Response.Write "</div>"
			End If
		End Function
	   			   
		Private Function RenderHTMLContent()
			Dim i
			Dim cn
			Dim fnc
			'text
			If LayerTemplate = "" Then
				If mHTML <> "" Then
					Response.Write mHTML
				End If
			Else
				On Error Resume Next
				Set fnc = GetRef(LayerTemplate)	
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