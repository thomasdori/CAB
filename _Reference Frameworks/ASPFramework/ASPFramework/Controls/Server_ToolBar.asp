<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server ToolBar

	Public Function New_ServerToolBar(name)
		Set New_ServerToolBar = New ServerToolBar
			New_ServerToolBar.Control.Name = name
	End Function

	Page.RegisterLibrary "ServerToolBar"

	Page.RegisterClientstartupScript "CLASP_ToolBarStyle", "<style> " + vbNewLine + _
		"	.tbButtonNormal { " + vbNewLine + _
		"		text-align:center; " + vbNewLine + _
		"		border:1px solid #efedde; " + vbNewLine + _
		"		color:black; " + vbNewLine + _
		"		cursor:default; " + vbNewLine + _
		"	} " + vbNewLine + _
		"	.tbButtonHover { " + vbNewLine + _
		"		text-align:center; " + vbNewLine + _
		"		color:black; " + vbNewLine + _
		"		border:1px solid navy; " + vbNewLine + _
		"		background-color:#abbee8; " + vbNewLine + _
		"		cursor:default; " + vbNewLine + _
		"	} " + vbNewLine + _
		"	.tbButtonSelect { " + vbNewLine + _
		"		text-align:center; " + vbNewLine + _
		"		border:1px solid navy; " + vbNewLine + _
		"		color:white; " + vbNewLine + _
		"		background-color:navy; " + vbNewLine + _
		"		cursor:default; " + vbNewLine + _
		"	} " + vbNewLine + _
		"</style> " + vbNewLine

	CONST tbHorizontal	= 1
	CONST tbVertical 	= 2
	
	Class ServerToolBar
		Public Control
		Public Layout
		Public Spacing
		Public BackColor
		
		Public ButtonStyle
		Public ButtonCssClass
		Public ButtonHoverCssClass
		Public ButtonSelectCssClass
		
		Private mRoot
		Private mItems
		
		Private Sub Class_Initialize()
			Dim Att	
			Dim mItem
			
			Set Control = New WebControl
			Set Control.Owner = Me

			Spacing					= 1
			Layout					= tbHorizontal
			BackColor				= "EFEBD7"
			ButtonStyle				= ""
			ButtonCssClass			= "tbButtonNormal"
			ButtonHoverCssClass		= "tbButtonHover"
			ButtonSelectCssClass	= "tbButtonSelect"

			Set mRoot = CreateObject("MSXML2.FreeThreadedDOMDocument.2.6")
			Set mItems = mRoot.createElement("I")
			mRoot.appendChild mItems

			Set mItem = mRoot.createElement("I")
			mItems.appendChild mItem
			
			With mItem
				Set Att = .ownerDocument.createAttribute("N")	'name/id
				.Attributes.setNamedItem Att
				
				Set Att = .ownerDocument.createAttribute("V")	'value
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("S")	'size
				Att.Text = "0"
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("T")	'tip
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("M")	'mode: 1 - button, 2 - buttonTemplate, 3 - separator
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("O")	'toggle state
				Att.Text = "0"
				.Attributes.setNamedItem Att

			End With
			
		End Sub
	   
		Public Function ReadProperties(bag)
			Layout	 			= CInt(bag.Read("L"))			
			Spacing	 			= CInt(bag.Read("S"))
			BackColor			= bag.Read("B")
			ButtonStyle		 	= bag.Read("Y")
			ButtonCssClass	 	= bag.Read("C")
			ButtonHoverCssClass	= bag.Read("H")
			ButtonSelectCssClass= bag.Read("E")
			mRoot.loadXML bag.Read("X")
			Set mItems = mRoot.firstChild
		End Function

		Public Function WriteProperties(bag)
			bag.Write "L",Layout
			bag.Write "S",Spacing
			bag.Write "B",BackColor
			bag.Write "Y",ButtonStyle
			bag.Write "C",ButtonCssClass
			bag.Write "H",ButtonHoverCssClass
			bag.Write "E",ButtonSelectCssClass
			bag.Write "X",mRoot.xml
		End Function
	
		Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunctionEX(e)
		End Function

		Public Function AddButton(Id,Value,Size,Tip)
			Dim nItem
			
			Set nItem = mItems.firstChild.CloneNode(0)
			mItems.appendChild nItem
			
			With nItem
				.Attributes.getNamedItem("N").Text = Id
				.Attributes.getNamedItem("V").Text = Value
				.Attributes.getNamedItem("S").Text = Size
				.Attributes.getNamedItem("T").Text = Tip
				.Attributes.getNamedItem("M").Text = 1	'Button
			End With			
			
		End Function

		Public Function AddButtonTemplate(Name,Size,Tip)
			Dim nItem			
			Call AddButton("",Name,Size,Tip)
			Set nItem = mItems.lastChild
			nItem.Attributes.getNamedItem("M").Text = 2	'ButtonTemplate Mode -  A Function name that is used to render HTML inside a button
		End Function
		
		Public Function AddSeparator()
			Dim nItem			
			Call AddButton("","",0,"")
			Set nItem = mItems.lastChild
			nItem.Attributes.getNamedItem("M").Text = 3	'Separator
		End Function

		Public Function ToggleButton(Id,State)
			Dim n
			Set n = mRoot.selectSingleNode("//I[@N='"+id+"']")
			n.Attributes.getNamedItem("O").Text = IIf(State,1,0)	'Toggle
		End Function

		Public Function ToggleAllButton(State)
			Dim i
			Dim n
			Set n = mRoot.selectNodes("//I")
			For i = 1 to n.Length -1
				n(i).Attributes.getNamedItem("O").Text = IIf(State,1,0)	'Toggle
			Next
		End Function

		Public Function GetButtonToggleState(Id)
			Dim n
			Set n = mRoot.selectSingleNode("//I[@N='"+Id+"']")
			If Not n Is Nothing Then 
				GetButtonToggleState = CBool(n.Attributes.getNamedItem("O").Text) 'Toggle
			End If
		End Function


		Private Function RenderToolBar()
			Dim i
			Dim Style
			Dim Fnc
			Dim N,V,M,T,S,O
			
			Response.Write "<table id='"+Control.ControlID+"' border='0' cellpadding='0' cellspacing='" & Spacing &"' bgcolor='" + BackColor + "' "
				If Control.CssClass <> ""	Then Response.Write " class='" & Control.CssClass  & "' "
				If Control.Style <> ""		Then Response.Write " style='" & Control.Style	  & "' "
				If Control.Attributes <> ""	Then Response.Write Control.Attributes   & " "
				Response.Write ">"				
				For i = 1 to mItems.childNodes.Length-1
					With mItems.childNodes(i)
						N = mItems.childNodes(i).Attributes.getNamedItem("N").Text
						V = .Attributes.getNamedItem("V").Text
						M = CInt(.Attributes.getNamedItem("M").Text)
						T = .Attributes.getNamedItem("T").Text
						S = CInt(.Attributes.getNamedItem("S").Text)
						O = CBool(.Attributes.getNamedItem("O").Text)
					End With
					Style = ButtonStyle & IIf(S>0,";width:" & S & ";","")
					If Layout = tbVertical Then
						Response.Write "<tr><td>"						
					Else
						If i=1 Then 
							Response.Write "<tr><td>"
						Else
							Response.Write "<td height='100%'>"
						End If
					End If
					Response.Write "<div "
					If T<>"" Then Response.Write " title='"+ Server.HTMLEncode(T) +"' "
					If M=1 And Style<>"" Then Response.Write "style='" + Style +"' "
					If M=1 And ButtonCssClass<>"" Then  Response.Write "class='" + IIf(O And ButtonSelectCssClass<>"",ButtonSelectCssClass, ButtonCssClass) +"' onmouseout='this.className="""+ IIf(O And ButtonSelectCssClass<>"",ButtonSelectCssClass, ButtonCssClass) +""";' "
					If M=1 And O=0 And ButtonHoverCssClass<>"" Then  Response.Write "onmouseover='this.className="""+ButtonHoverCssClass+""";' "
					If M=1 Then Response.Write Page.GetEventScript("onclick",Control.ControlID,"OnClick",N,"") +" "
					Response.Write ">"	
						Select Case M
							Case 1:
								Response.Write V
							Case 2:
								Set Fnc = GetFunctionReference(v)
								If Not Fnc is Nothing Then Fnc()
							Case 3:
								If Layout = tbHorizontal Then
									Response.Write "<table height='100%' border='0' cellpading='0' cellspacing='1'><tr><td height='100%' background='"& SCRIPT_LIBRARY_PATH &"toolbar/images/vsep.gif'></td></tr></table>"
								Else
									Response.Write "<table width='100%' border='0' cellpading='0' cellspacing='1'><tr><td height='2' width='100%' background='"& SCRIPT_LIBRARY_PATH &"toolbar/images/hsep.gif'></td></tr></table>"
								End If
						End Select
					Response.Write "</div>"
					Response.Write "</td>"
					If Layout = tbVertical Then Response.Write "</tr>"
				Next
				If Layout = tbHorizontal Then Response.Write "</tr>"
			Response.Write "</table>"
		End Function

		Public Default Function Render()
			Dim varStart

			If Control.IsVisible = False Then
				Exit Function
			End If

			varStart = Now
			Call RenderToolBar()
			Page.TraceRender varStart,Now,Control.Name
		End Function
		
	End Class

%>