<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server ViewPort

	Public Function New_ServerViewPort(name,url,width,height)
		Set New_ServerViewPort = New ServerViewPort
			New_ServerViewPort.Control.Name = name
			New_ServerViewPort.LocationURL	= url
			New_ServerViewPort.Width		= width
			New_ServerViewPort.Height		= height
	End Function

	Page.RegisterLibrary "ServerViewPort"

	Page.RegisterClientStartupScript "ResizeViewPort", _
		"<script> "+ vbNewLine + _
		"	function ResizeViewPort(n,mode){ " + vbNewLine + _
		"		if (!n && document.layers) return null;" + vbNewLine + _
		"		var f=(document.getElementById)? document.getElementById(n): document.all[n];" + vbNewLine + _ 
		"		var doc = (f && f.contentDocument)? f.contentDocument:f.Document;" + vbNewLine + _ 
		"		var inp = clasp.form.getField(n+'_URL');" + vbNewLine + _
		"		if (doc && inp) {" + vbNewLine + _
		"			 // save url for post backs " + vbNewLine + _
		"			inp.value = (doc.location.href!=document.location.href)? doc.location:'';" + vbNewLine + _
		"		}" + vbNewLine + _
		"		var c = eval(n+'_content');" + vbNewLine + _ 
		"		eval(n+'_content=null'); // reset content" + vbNewLine + _ 
		"		if (c && doc) {" + vbNewLine + _
		"			doc.open();" + vbNewLine + _
		"			doc.write(c);" + vbNewLine + _
		"			doc.close();" + vbNewLine + _
		"		}" + vbNewLine + _
		"		w = (doc.body.scrollWidth)? doc.body.scrollWidth:doc.body.offsetWidth;" + vbNewLine + _
		"		h = (doc.body.scrollHeight)? doc.body.scrollHeight:doc.body.offsetHeight;" + vbNewLine + _
		"		if(mode==1||mode==3) f.style.width = w;" + vbNewLine + _
		"		if(mode==2||mode==3) f.style.height = h;" + vbNewLine + _
		"		if (!f.firstLoad) {clasp.form.resetScroll();f.firstLoad=true;}" + vbNewLine + _
		"	}" + vbNewLine + _
		"</script>"

	Class ServerViewPort
		Public Control
		Public Width
		Public Height
		Public AutoSize			'resize the fit it's content's with and height
		Public AutoWidth		'resize the viewport to fit it's content's width
		Public AutoHeight		'resize the viewport to fit it's content's height
		Public StretchToFit		'stretches the viewport to fit the available space (both width and height)
		Public ScrollBarMode	'0 - No scrollbar, 1 - Auto, 2 - Always show bars
		Public BorderSize
		Public LocationURL
		Public InnerHTML		'used to insert html inside the viewport if LocationURL was not set
		
		Private Sub Class_Initialize()

			Set Control = New WebControl
			Set Control.Owner = Me
				Control.ImplementsProcessPostBack = True

			Width			= 100
			Height			= 100
			AutoSize		= False
			AutoWidth		= False
			AutoHeight		= False
			StretchToFit		= False
			ScrollBarMode	= 1	'Auto
			BorderSize		= 0
			LocationURL		= ""
			InnerHTML		= ""
		End Sub
	   
		Public Function ReadProperties(bag)
			Dim a
			Dim props

			props 		= bag.Read("P")
			LocationURL = bag.Read("U")
			InnerHTML 	= bag.Read("H")

			'get properties from array
			a = Split(props,";")
			If UBound(a)>0 Then
				Width			= a(0)
				Height			= a(1)
				AutoSize		= a(2)
				AutoWidth		= a(3)
				AutoHeight		= a(4)
				StretchToFit	= a(5)
				ScrollBarMode	= a(6)
				BorderSize		= a(7)
			End If
		End Function

		Public Function WriteProperties(bag)
			Dim a
			Dim props

			'store properties inside and array
			a = Array( _
					width, _
					height, _
					AutoSize, _
					AutoWidth, _
					AutoHeight, _
					StretchToFit, _
					ScrollBarMode, _
					BorderSize _
				)
			props = Join(a,";")

			bag.Write "P", props
			bag.Write "U",LocationURL
			bag.Write "H",InnerHTML
		End Function

		Public Function ProcessPostBack()
			Dim url
			url = Request.Form(Control.ControlID+"_URL")
			If url <> "" Then
				LocationURL  = url
			End If
		End Function
		
		Public Function SetValueFromDataSource(value)		
			LocationURL = value
		End Function

		Public Function HandleClientEvent(e)
			HandleClientEvent = False
		End Function

		Private Function RenderViewPort()
			Dim sTag
			Dim eTag
			Dim w,h,autoMode

			If Browser.IsNS And Browser.Version = 4 Then
				sTag = "<ilayer"
				eTag = "</ilayer>"
			Else
				sTag = "<iframe"
				eTag = "</iframe>"
			End If

			autoMode = "0"
			w = CStr(Width)
			h = CStr(Height)
			If StretchToFit Then
				w = "100%":h = "100%"
			End If
			If AutoSize  Or (AutoWidth And AutoHeight) Then
				autoMode = "3" 'both
			ElseIf AutoWidth Then
				autoMode = "1" 'width only
			ElseIf AutoHeight Then
				autoMode = "2"	'height only
			End If
	
			Response.write "<script> var "+Control.ControlID+"_content = '"+ IIf(LocationURL="",CJString(InnerHTML),"")+"';</script>"
			Response.write "<input name='"+Control.ControlID+"_URL' type='hidden' value='"+Server.HTMLEncode(LocationURL)+"'>"
			Response.Write sTag +" id='"+Control.ControlID+"' name='"+Control.ControlID+"' "
				If Control.CssClass <> ""	Then Response.Write " class='" & Control.CssClass  & "' "
				If Control.Style <> ""		Then Response.Write " style='" & Control.Style	  & "' "
				If Control.Attributes <> ""	Then Response.Write Control.Attributes   & " "
				Response.Write "scrolling='"& IIF(ScrollBarMode = 0,"no",IIF(ScrollBarMode = 1,"auto","yes")) &"' "
				Response.Write "src='"+LocationURL+"' frameborder='"+CStr(BorderSize)+"' "
				Response.Write "onload=""window.setTimeout('ResizeViewPort(\'"+Control.ControlID+"\',"& autoMode &")',100);"" width='"+w+"' height='"+h+"' "
				Response.Write ">"
			Response.Write eTag
		End Function

		Public Default Function Render()
			Dim varStart

			If Control.IsVisible = False Then
				Exit Function
			End If

			varStart = Now
			Call RenderViewPort()
			Page.TraceRender varStart,Now,Control.Name
		End Function

		'CJString - converts server-side string (multi-line) to javascript (signle-line) string
		Private Function CJString(text)
			if Len(text)=0 Then Exit Function
			text=Replace(text,"\","\\")			' replace \ with \\
			text=Replace(text,"'","\'")			' replace ' with \'
			text=Replace(text,vbCrLf,"\n")		' replace CrLf with \n
			text=Replace(text,vbLf,"\n")		' replace single Lf with \n
			text=Replace(text,vbCr,"\r")		' replace single Cr with \n
			text=Replace(text,"</script>","<\/script>",1,-1,1)	' replace js ending </script> tag with <\/script>
			CJString = text 
		End Function
		
	End Class

%>