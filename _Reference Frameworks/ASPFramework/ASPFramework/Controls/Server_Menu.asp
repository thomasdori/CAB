<%
	'http://www.softcomplex.com/products/tigra_menu/docs/
	Page.RegisterClientStartupScript "ServerMenu","<script language='javascript' src='" & SCRIPT_LIBRARY_PATH & "menu/menu.js'></script>"

	Public Function New_ServerMenu(name,cache) 
		Set New_ServerMenu = New ServerMenu
			New_ServerMenu.Control.Name = name
	End Function
		
	Page.RegisterLibrary "ServerMenu"
		
	Class ServerMenuItem
	
		Public MenuNode 'As MSXML2.IXMLDOMNode
	
		Public Function SetAttribute(Name,Value)
			Dim node
			Set node =  MenuNode.Attributes.GetNamedItem(Name)
			If  node Is Nothing Then
				Set node  = MenuNode.OwnerDocument.CreateAttribute(Name)
				MenuNode.Attributes.SetNamedItem node
			End If
			node.Text = Value
			Set node = Nothing					
		End Function
		
		Public Function GetAttribute(Name)
			Dim node
			Set node =  MenuNode.Attributes.GetNamedItem(Name)
			If node Is Nothing Then
				GetAttribute = ""
			Else
				GetAttribute = node.Text
			End If
			Set node = Nothing
		End Function
		
		Public Property Get Text()
			Text = MenuNode.Attributes.getNamedItem("T").Text
		End Property
		
		Public Property Let Text(s_text)
			MenuNode.Attributes.getNamedItem("T").Text = s_text
		End Property
		
		Public Property Get Url()
			Url = MenuNode.Attributes.getNamedItem("U").Text
		End Property
		
		Public Property Let Url(s_url)
			MenuNode.Attributes.getNamedItem("U").Text = s_url
		End Property
		
		Public Property Get ItemScope()
			ItemScope = MenuNode.Attributes.getNamedItem("S").Text
		End Property
		
		Public Property Let ItemScope(s_item_scope)
			MenuNode.Attributes.getNamedItem("S").Text = s_item_scope
		End Property
		
		Public Property Get Visible()
			Visible = CBool(MenuNode.Attributes.getNamedItem("V").Text)		    
		End Property
		
		Public Property Let Visible(b_visible)
			MenuNode.Attributes.getNamedItem("V").Text = b_visible
		End Property
		
		Public Property Get Items(index)
			Dim oItem		    
			Set oItem = New ServerMenuItem
			Set oItem.MenuNode = MenuNode.childNodes(index)
			Set Items = oItem
		End Property

		Public Function Add(Text, Url, ItemScope, SubItems)
			Dim Item 'As MSXML2.IXMLDOMNode
			Dim MenuItem 'As New ServerMenuItem
			Dim MenuSubItem
			
			Set MenuItem  = New ServerMenuItem
			Set Item = MenuNode.OwnerDocument.firstchild.cloneNode(False)
			With Item.Attributes
				.getNamedItem("T").Text = Text
				.getNamedItem("U").Text = Url
				.getNamedItem("S").Text = ItemScope
		    End With
			MenuNode.appendChild item
			Set MenuItem.MenuNode = item
		    
			If IsArray(SubItems) Then
				For Each MenuSubItem In SubItems
					MenuItem.Add MenuSubItem(0), MenuSubItem(1), MenuSubItem(2), ""
				Next
			End If
		    
			Set Add = MenuItem
		    
		End Function
		
		Public Function AddItems(SubItems)
			Dim MenuSubItem
			If IsArray(SubItems) Then
				For Each MenuSubItem In SubItems
					Add MenuSubItem(0), MenuSubItem(1), MenuSubItem(2), ""
				Next
			End If
		End Function
		
		Public Function AddMenuItem(MenuItem)
			AddItem MenuItem.Text, MenuItem.Url, MenuItem.ItemScope, ""
		End Function
		
		Public Function Remove(v)
			MenuNode.removeChild MenuNode.childNodes(v)
		End Function
		
		Public Property Get Count()
			Count = MenuNode.childNodes.length
		End Property

		Public Property Get Parent
			Dim Node
			Set Node = MenuNode.ParentNode
			If Node Is Nothing Then
				Set Parent = Nothing
			Else
				Set Parent = New ServerMenuItem
				Set Parent.MenuNode = Node
			End If	
		End Property
		
		Public Function GetJs(buffer, lvl)
			Dim x, mx
			Dim mItem
		                            
			buffer.Append vbNewLine & String(lvl, vbTab) & "[""" & Text + """," & IIf(Url <> "", """" & Url & """", "null") & "," & IIf(ItemScope <> "", "{" & ItemScope & "}", "null") & ","
			mx = Count - 1
				If mx >= 0 Then
					For x = 0 To mx
						Set mItem = Items(x)
						If mItem.Visible Then
							mItem.GetJs buffer, lvl + 1
							buffer.Append "," '& vbNewLine
						End If
					Next
					buffer.Append vbNewLine & String(lvl, vbTab) & "]"
				Else
					buffer.Append "null]"
				End If
		End Function
	
	End Class
	
':::::::::::::::::
	Class ServerMenu
		Private mobjXMLDocument 'As MSXML2.DOMDocument26
		Private MenuNode 'As MSXML2.IXMLDOMNode
		Private MnHome 'As ServerMenuItem
	
		Dim Control
		Dim MenuItemScope
		Dim ItemScopeName
		Dim mCacheName
						
		Dim Left
		Dim Top
			
		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me				

			MenuItemScope = "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "menu/item_scope.js'></script>"
			ItemScopeName = "MENU_POS"
			
			Dim Att 'As MSXML2.IXMLDOMNode
			Set mobjXMLDocument = CreateObject("MSXML2.FreeThreadedDOMDocument.2.6")			
			Set MenuNode = mobjXMLDocument.createElement("I")			
			mobjXMLDocument.appendChild MenuNode
			With MenuNode
				Set Att = .ownerDocument.createAttribute("T")
				.Attributes.setNamedItem Att
				Set Att = .ownerDocument.createAttribute("U")
				.Attributes.setNamedItem Att
				Set Att = .ownerDocument.createAttribute("S")
				.Attributes.setNamedItem Att
				Set Att = .ownerDocument.createAttribute("V")
				Att.Text = "1"
				.Attributes.setNamedItem Att
			End With
			Set MnHome = New ServerMenuItem
			Set MnHome.MenuNode = MenuNode
		End Sub
		
		Private Sub Class_Terminate()
			Set MenuNode = Nothing
			Set mobjXMLDocument = Nothing
			Set MnHome = Nothing
		End Sub

		Public Property Get CacheName
			CacheName = mCacheName
		End Property

		Public Property Let CacheName(s_cache)
			mCacheName = s_cache
			If Not Page.IsPostBack And s_cache <> "" Then
				If mobjXMLDocument.loadXML(Application(s_cache)) Then'
					Set MnHome.MenuNode = mobjXMLDocument.firstChild
				End If
			End If
		End Property

		Public Function ReadProperties(bag)			
			If mCacheName = "" Then
				mobjXMLDocument.loadXML bag.Read("MENU")
			Else
				mobjXMLDocument.loadXML(Application(mCacheName))
			End If
			Set MnHome.MenuNode = mobjXMLDocument.firstChild
		End Function
		
			
		Public Function WriteProperties(bag)
			bag.Write "MENU",mobjXMLDocument.xml
		End Function

		Public Function HandleClientEvent(e)
			If e.Instance = "0" Then
				Dim menu
				Dim fnc
				Set fnc = GetFunctionReference(Control.ControlID & "_OnClick")
				If Not fnc Is Nothing Then
					Set menu = GetClickedMenu(CInt(e.ExtraMessage))
					Call fnc(menu)					
				End If
			Else
				HandleClientEvent = ExecuteEventFunctionEX(e)
			End If
		End Function			

		Public Function GetMenuItem(sTree) 'As ServerMenuItem
			Dim vTree
			Dim MenuIndex
			Dim MenuItem 'As ServerMenuItem
		    
			vTree = Split(sTree, ",")
			Set MenuItem = MnHome.Item(vTree(0))
			For MenuIndex = 1 To UBound(vTree)
				Set MenuItem = MenuItem.Item(vTree(MenuIndex))
			Next		    
		End Function

		Public Property Get Count
			Count = MnHome.Count
		End Property
		
		Public Function Add(Text, Url, ItemScope, SubItems)
			Set Add = MnHome.Add(Text, Url, ItemScope, SubItems)
		End Function

		Public Function Items(index) 'As ServerMenuItem
			Set Items = MnHome.Items(index)
		End Function

		Public Function Remove(index)
			MnHome.Remove index
		End Function

		Public Function GetXml()
			GetXml = mobjXMLDocument.xml
		End Function

		Public Default Function Render()

			'new menu (MENU_ITEMS, MENU_POS);
			Dim buffer
			Dim mitem
			Dim x, mx
			Dim varStart
		        
			If Control.IsVisible = False Then
				Exit Function
			End If

			varStart = Now

			Response.Write vbNewLine & MenuItemScope & vbNewLine
			
			Set buffer = New StringBuilder
			buffer.GrowRate = 100
			mx = MnHome.Count - 1
			buffer.Append "var " & Control.ControlID & "_MENU_ITEMS = [ "
			
			For x = 0 To mx
				Set mitem = MnHome.Items(x)
				If mitem.Visible Then
					mitem.GetJs buffer, 1
					buffer.Append ","
				End If
			Next
			buffer.Append "];"

			Response.Write vbNewLine & "<Script Language='JavaScript'>" & vbNewLine
				If Left <> "" Then Response.Write ItemScopeName & "[0].block_left= " & Left & ";" & vbNewLine
				If Top<> "" Then Response.Write ItemScopeName & "[0].block_top= " & Top & ";" & vbNewLine					
				Response.Write buffer.ToString()
				Response.Write vbNewLine & " new menu(" & Control.ControlID & "_MENU_ITEMS," & ItemScopeName & ",'" + Control.ControlID + "');" & vbNewLine
			Response.Write vbNewLine & "</Script>" & vbNewLine
			Page.TraceRender varStart, Now, Control.ControlID
			
		End Function

		Public Function Find(MenuText)
			Dim oNode 
			Dim MenuItem
			Set oNode = mobjXMLDocument.selectSingleNode("//I[@T='" & Replace(MenuText,"'","''") &  "']")
			If oNode Is Nothing Then
			   Set Find = Nothing
			Else
			   Set MenuItem = New ServerMenuItem
			   Set MenuItem.MenuNode = oNode
			   Set Find = MenuItem
		    End If						
		End Function
		Public Function FindMenu(PropName,PropValue)
			Dim oNode 
			Dim MenuItem
			Set oNode = mobjXMLDocument.selectSingleNode("//I[@" & PropName &  "='" & Replace(PropValue,"'","''") &  "']")
			If oNode Is Nothing Then
			   Set FindMenu = Nothing
			Else
			   Set MenuItem = New ServerMenuItem
			   Set MenuItem.MenuNode = oNode
			   Set FindMenu = MenuItem
		    End If						
		End Function
		
		Private Function GetClickedMenu(i)
			Dim MenuItem
			Dim oList

			Set MenuItem = New ServerMenuItem
			Set oList = mobjXMLDocument.selectNodes("//")
			Set MenuItem.MenuNode = oList(i+1)
			Set oList = Nothing
			Set GetClickedMenu = MenuItem
		End Function


	End Class		
%>