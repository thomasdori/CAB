<%  'http://www.softcomplex.com/products/tigra_tree_menu/docs/
	Page.RegisterClientStartupScript "ServerTreeView","<script language='javascript' src='" & SCRIPT_LIBRARY_PATH & "treeview/tree.js'></script>"
	Page.RegisterLibrary "ServerTreeView"		

	Public Function New_ServerTreeView(name,cache)
		Set New_ServerTreeView = New ServerTreeView
			New_ServerTreeView.Control.Name = name
			If cache<>"" Then
				New_ServerTreeView.TreeCacheName = cache
			End If
	End Function		

	Class ServerTreeViewNode
		'Text,Url OR 0, SubItems	
		Public TreeNode 'As MSXML2.IXMLDOMNode

		Public Function SetAttribute(Name,Value)
			Dim node
			Set node =  TreeNode.Attributes.GetNamedItem(Name)
			If  node Is Nothing Then
				Set node  = TreeNode.OwnerDocument.CreateAttribute(Name)
				TreeNode.Attributes.SetNamedItem node
			End If
			node.Text = Value
			Set node = Nothing					
		End Function
		
		Public Function GetAttribute(Name)
			Dim node
			Set node =  TreeNode.Attributes.GetNamedItem(Name)
			If node Is Nothing Then
				GetAttribute = ""
			Else
				GetAttribute = node.Text
			End If
			Set node = Nothing
		End Function


		Public Property Get Text()
			Text = TreeNode.Attributes.getNamedItem("T").Text
		End Property
		
		Public Property Let Text(s_text)
			TreeNode.Attributes.getNamedItem("T").Text = s_text
		End Property
		
		Public Property Get Url()
			Url = TreeNode.Attributes.getNamedItem("U").Text
		End Property
		
		Public Property Let Url(s_url)
			TreeNode.Attributes.getNamedItem("U").Text = s_url
		End Property
		
		Public Property Get IsOpened()
			IsOpened = CBool(TreeNode.Attributes.getNamedItem("O").Text)
		End Property
		
		Public Property Let IsOpened(b_opened)
			TreeNode.Attributes.getNamedItem("O").Text = IIF(b_opened,1,0)
		End Property
		
		Public Property Get Visible()
			Visible = CBool(TreeNode.Attributes.getNamedItem("V").Text)		    
		End Property
		
		Public Property Let Visible(b_visible)
			TreeNode.Attributes.getNamedItem("V").Text = b_visible
		End Property
		
		Public Property Get Nodes(index) 'As ServerTreeViewNode
			Dim oItem		    
			Set oItem = New ServerTreeViewNode
			Set oItem.TreeNode = TreeNode.childNodes(index)
			Set Nodes = oItem
		End Property

		Public Function AddNode(Text, Url, Nodes)
			Dim NewNodeXML 'As MSXML2.IXMLDOMNode
			Dim NewNode 'As New ServerTreeViewNode
			Dim Node
			
			Set NewNode    = New ServerTreeViewNode
			'Set NewNodeXML = TreeNode.cloneNode(False)
			Set NewNodeXML = TreeNode.ownerDocument.childNodes(0).CloneNode(0)
			
			NewNodeXML.Attributes.getNamedItem("T").Text = Text
			NewNodeXML.Attributes.getNamedItem("U").Text = Url
		    
			TreeNode.appendChild NewNodeXML
			
			Set NewNode.TreeNode = NewNodeXML
		    
			If IsArray(Nodes) Then
				For Each Node In Nodes
					NewNode.AddNode Node(0), Node(1), ""
				Next
			End If
		    
			Set AddNode = NewNode
		    
		End Function
		
		Public Function AddNodes(Nodes)
			Dim Node
			If IsArray(Nodes) Then
				For Each Node In Nodes
					Me.AddNode Node(0), Node(1), ""
				Next
			End If
		End Function
				
		Public Function Remove(index)
			TreeNode.removeChild TreeNode.childNodes(index)
		End Function
		
		Public Property Get Count()
			Count = TreeNode.childNodes.length
		End Property

		Public Property Get Parent
			Dim Node
			Set Node = TreeNode.ParentNode
			If Node Is Nothing Then
				Set Parent = Nothing
			Else
				Set Parent = New ServerTreeViewNode
				Set Parent.TreeNode = Node
			End If	
		End Property
		
		Public Function GetJs(buffer, lvl)
			Dim x, mx
			Dim mitem
		                            
			buffer.Append vbNewLine & String(lvl, vbTab) & "[""" & Text + """," & IIf(Url <> "", """" & Url & """", "null") & "," 
			mx = Count - 1
				If mx >= 0 Then
					For x = 0 To mx
						Set mitem = Nodes(x)
						If mitem.Visible Then
							mitem.GetJs buffer, lvl + 1
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
	Class ServerTreeView
		Private mobjXMLDocument 'As MSXML2.DOMDocument26
		Private TreeNode 'As MSXML2.IXMLDOMNode
		Private tvRoot 'As ServerTreeViewNode

		Dim Control
		Dim MenuStyle 'Path to the item scope file
		Dim MenuStyleName 'Name of the Js item scope variable				
		Dim mTreeCacheName
		
		Private sTreeOpenList
		Private SelectedNodeIndex
			
		Private Sub Class_Initialize()
			Dim Att
			
			Set Control = New WebControl	
			Set Control.Owner = Me				
			
			Set mobjXMLDocument = CreateObject(CLASP_DOM_PARSER_CLASS)			
			
			Set TreeNode = mobjXMLDocument.createElement("I")			
			mobjXMLDocument.appendChild TreeNode
			
			With TreeNode
				Set Att = .ownerDocument.createAttribute("T")
				.Attributes.setNamedItem Att
				
				Set Att = .ownerDocument.createAttribute("U")
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("O")
				Att.Text = "0"
				.Attributes.setNamedItem Att

				Set Att = .ownerDocument.createAttribute("V")
				Att.Text = "1"
				.Attributes.setNamedItem Att
				
			End With
						
			Set tvRoot = New ServerTreeViewNode
			Set tvRoot.TreeNode = TreeNode
			
			MenuStyle = "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "treeview/default.js'></script>"
			MenuStyleName = "tvTree_Default"
		End Sub
		
		Private Sub Class_Terminate()
			Set TreeNode = Nothing
			Set mobjXMLDocument = Nothing
			Set tvRoot = Nothing
		End Sub
		
		Public Property Get TreeCacheName
			TreeCacheName = mTreeCacheName
		End Property

		Public Property Let TreeCacheName(s_cache)		
			mTreeCacheName = s_cache
			If Not Page.IsPostBack And s_cache <> "" Then
				If mobjXMLDocument.loadXML(Application(s_cache)) Then
					Set tvRoot.TreeNode = mobjXMLDocument.firstChild
				End If
			End If
		End Property
		
		Public Function ReadProperties(bag)			
			Dim vOpenedItems
			Dim oList
			Dim x,mx
			Dim bUpdateTree						
			bUpdateTree = False			
		
			If mTreeCacheName = "" Then
				mobjXMLDocument.loadXML bag.Read("TV")
				Set tvRoot.TreeNode = mobjXMLDocument.firstChild
			Else
				If mobjXMLDocument.loadXML(Application(mTreeCacheName)) Then
					Set tvRoot.TreeNode = mobjXMLDocument.firstChild
				Else
					Response.Write "Could Not Load TreeView " & Control.ControlID
					With mobjXMLDocument
						Response.Write "Reason:" & .parseError.reason
						Response.Write "<br>errorCode:" & .parseError.errorCode
						Response.Write "<br>srcText:[" & .parseError.srcText & "]"
						Response.Write "<br>Line:" & .parseError.line
						Response.Write "<br>Col:" & .parseError.linepos
						Response.End
					End With			
				End If
				bUpdateTree = True
			End If
			
			sTreeOpenList = Request.Form(Control.ControlID & "_Status")			
			SelectedNodeIndex = bag.Read("S")
			
			If sTreeOpenList <> bag.Read("O") Then
				bUpdateTree  = True
			End If
			If bUpdateTree  Then			
				vOpenedItems = Split(sTreeOpenList,",")

				Set oList = mobjXMLDocument.selectNodes("//I[@O='1']")
				mx  = olist.length-1
				For x = 1 To mx
					oList(x).Attributes.getNamedItem("O").Text = 0
				Next
			
				Set oList = mobjXMLDocument.selectNodes("//")
				mx  = UBound(vOpenedItems)
				For x = 0 To mx
					If vOpenedItems(x)<>"" Then
						oList(CLng(vOpenedItems(x)+1)).Attributes.getNamedItem("O").Text = 1
					End If
				Next
			End If			

		End Function
		
			
		Public Function WriteProperties(bag)
			If mTreeCacheName = "" Then
				bag.Write "TV",mobjXMLDocument.xml
			End If			
		End Function

		Public Function HandleClientEvent(e)						
			Dim node
			Dim fnc
			HandleClientEvent = True
		
			SelectedNodeIndex = e.Instance
			Set fnc = GetFunctionReference(Control.ControlID & "_" & e.EventName )			
			If Not fnc Is Nothing Then
				Set node = GetClickedNode(CInt(SelectedNodeIndex))
				Call fnc(node)					
			End If
		End Function			

		Public Property Get Count
			Count = tvRoot.Count
		End Property
		
		Public Function AddRoot(Text, Url, SubItems)
			Set AddRoot = tvRoot.AddNode(Text, Url, SubItems)
		End Function

		Public Function Nodes(index) 'As ServerTreeViewNode
			Set Nodes = tvRoot.Nodes(index)
		End Function

		Public Function Remove(index)
			tvRoot.Remove index
		End Function

		Public Default Function Render()

			'tree (TREE_ITEMS, TREE_TPL)
			Dim buffer
			Dim mitem
			Dim x, mx
			Dim varStart
			Dim TreeImages
						
			If Control.IsVisible = False Then
				Exit Function
			End If
						
			varStart = Now
			
			Response.Write MenuStyle
			Set buffer = New StringBuilder
			buffer.GrowRate = 100
			mx = tvRoot.Count - 1
			buffer.Append "var " & Control.ControlID & "_NODES = [ "
			
			For x = 0 To mx
				Set mitem = tvRoot.Nodes(x)
				If mitem.Visible Then
					mitem.GetJs buffer, 1
					buffer.Append ","
				End If
			Next
			buffer.Append "];"
			
			sTreeOpenList = GetSelectedList()
			Response.Write "<input type='hidden' name='" & Control.ControlID & "_Status'>"
			Response.Write vbNewLine & "<Script Language='JavaScript'>" & vbNewLine
			Response.Write buffer.ToString()
			Response.Write vbNewLine & " new tree(" & Control.ControlID & "_NODES," & MenuStyleName & ",'" & Control.ControlID + "','" & sTreeOpenList & "'," & IIF(SelectedNodeIndex<>"",SelectedNodeIndex,"null") & ")"
			Response.Write vbNewLine & "</Script>" & vbNewLine		
			
			If Not Control.Viewstate Is Nothing Then
				Control.ViewState.Write "S",SelectedNodeIndex
				Control.ViewState.Write "O",sTreeOpenList
			End If
			
			Page.TraceRender varStart, Now, Control.ControlID
						
		End Function

		Private Function GetTreeNode(node, i)
			Dim x, mx, imx
			Set GetTreeNode = Nothing
		    
			If i = 0 Then
				Set GetTreeNode = node
			Else
				mx = node.childNodes.length - 1
		        
				If mx >= 0 Then
					For x = 0 To mx
		                
						i = i - 1
						If i = 0 Then
							Set GetTreeNode = node.childNodes(x)
							Exit For
						End If
		                
						If node.childNodes(x).childNodes.length > 0 Then
							Set GetTreeNode = GetTreeNode(node.childNodes(x), i)
							If Not GetTreeNode Is Nothing Then
								Exit For
							End If
						End If
		            
					Next
				End If
			End If
		    
		End Function

		Private Function GetSelectedList()
			Dim oList
			Dim x,mx
			Dim sSelected
			Set oList = mobjXMLDocument.selectNodes("//")
			mx = oList.Length - 1
			For x = 1 To mx
				If oList(x).Attributes.GetNamedItem("O").Text = "1" Then
					sSelected = sSelected & (x-1) & ","
				End If
			Next
			GetSelectedList = sSelected
		End Function
	
		Private Function GetClickedNode(i)
			Dim MenuItem
			Set MenuItem = New ServerTreeViewNode
			Set MenuItem.TreeNode = GetTreeNode(mobjXMLDocument.firstChild, i+1)
			Set GetClickedNode = MenuItem
		End Function
		
		Public Function GetXml()
			GetXml = mobjXMLDocument.xml
		End Function
		
		Public Function LoadXml(sXml)
			LoadXml = mobjXMLDocument.loadXML(sXml)
		End Function
		
		Public Function LoadFromFile(sFile)
			LoadFromFile = mobjXMLDocument.load(sFile)
		End Function
		
		Public Function SaveToFile(sFile)
			mobjXMLDocument.save sFile
		End Function

	End Class	
	
%>