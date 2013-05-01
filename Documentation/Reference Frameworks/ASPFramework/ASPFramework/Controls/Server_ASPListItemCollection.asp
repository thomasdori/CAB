<%

	Class cASPListItemCollection
		Private mobjXMLDocument
		Private mobjItemNode
		Private mMode

		Private Sub Class_Initialize()
			'Initializes the ItemCollection

			'Set mobjXMLDocument = Server.CreateObject("MSXML2.FreeThreadedDOMDocument.3.0")
			Set mobjXMLDocument = Server.CreateObject(CLASP_DOM_PARSER_CLASS)

			mobjXMLDocument.loadXML "<IC><I V="""" S="""" T=""""/></IC>"

			Set mobjItemNode = mobjXMLDocument.firstChild.childNodes(0).cloneNode(False)

			mobjXMLDocument.firstChild.removeChild mobjXMLDocument.firstChild.childNodes(0)

			mMode = 0 'Single Select

		End Sub

		Public Property Get GetSelectedText()
			Dim strCritera 
			Dim objNode

			strCritera = "//IC/I[@S='-1']"

			Set objNode = mobjXMLDocument.selectSingleNode(strCritera)

			If objNode Is Nothing Then
				GetSelectedText = ""
			Else
				GetSelectedText = objNode.Attributes.getNamedItem("T").nodeValue
			End If

		End Property

		Public Property Get GetSelectedValue()
			Dim strCritera 
			Dim objNode

			strCritera = "//IC/I[@S='-1']"

			Set objNode = mobjXMLDocument.selectSingleNode(strCritera)

			If objNode Is Nothing Then
				GetSelectedValue = ""
			Else
				GetSelectedValue = objNode.Attributes.getNamedItem("V").nodeValue
			End If

		End Property

		Public Property Let Mode(ByVal value)
			mMode = value
		End Property

		Public Property Get Mode()
			Mode = mMode
		End Property

		Public Sub SetState(ByRef strCollectionXML)
				mobjXMLDocument.loadXML strCollectionXML
		End Sub

		Public Function GetState()
			GetState = mobjXMLDocument.xml
		End Function

		Public Sub Clear()
			Set mobjXMLDocument = Server.CreateObject(CLASP_DOM_PARSER_CLASS)

			mobjXMLDocument.loadXML "<IC><I V="""" S="""" T=""""/></IC>"

			Set mobjItemNode = mobjXMLDocument.firstChild.childNodes(0).cloneNode(False)

			mobjXMLDocument.firstChild.removeChild mobjXMLDocument.firstChild.childNodes(0)

		End Sub

		Public Sub Add(ByVal Text , ByVal value , ByVal selected , ByVal index)
			Dim xmlNode
			Dim xmlRef

			If index >= 0 Then
				If mobjXMLDocument.firstChild.childNodes.length < index Then
					index = mobjXMLDocument.firstChild.childNodes.length
				End If

				Set xmlRef = mobjXMLDocument.firstChild.childNodes(index)
			Else
				Set xmlRef = Nothing
			End If

			If mMode = 0 And selected = True Then
				Me.SetAllSelected False
			End If

			With mobjXMLDocument.firstChild.insertBefore(mobjItemNode.cloneNode(False), xmlRef)
				.Attributes.getNamedItem("T").nodeValue = Text
				.Attributes.getNamedItem("V").nodeValue = value
				.Attributes.getNamedItem("S").nodeValue = selected
			End With

		End Sub

		Public Sub Append(ByVal Text , ByVal value , ByVal selected)
			Me.Add Text,Value,Selected,-1
		End Sub

		Public Property Get Count() 
			Count = mobjXMLDocument.firstChild.childNodes.length
		End Property

		Public Sub Remove(ByVal index)
			If index >= 0 And index < Count Then
				mobjXMLDocument.firstChild.removeChild mobjXMLDocument.firstChild.childNodes(index)
			End If
		End Sub

		Public Function GetValue(ByVal index )
			If index >= 0 And index < Count Then
				GetValue = mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("V").nodeValue
			End If
		End Function

		Public Function GetText(ByVal index )
			If index >= 0 And index < Count Then
				GetText = mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("T").nodeValue
			End If
		End Function

		Public Function IsSelected(ByVal index)
			IsSelected = False
			If index >= 0 And index < Count Then
				IsSelected = mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("S").nodeValue
			End If
		End Function

		Public Sub GetItemData(ByVal index , ByRef Text , ByRef value , ByRef selected )
			Dim objNode
			If index >= 0 And index < Count Then
				Set objNode = mobjXMLDocument.firstChild.childNodes(index)
				selected = objNode.Attributes.getNamedItem("S").nodeValue
				Text = objNode.Attributes.getNamedItem("T").nodeValue
				value = objNode.Attributes.getNamedItem("V").nodeValue
			End If
		End Sub

		Public Function IsSelectedByText(ByVal Text) 
			Dim objNode

			Set objNode = GetItemByText(Text)

			If Not objNode Is Nothing Then
				IsSelectedByText = objNode.Attributes.getNamedItem("S").nodeValue
			End If

		End Function

		Public Function IsSelectedByValue(ByVal value )
			Dim objNode

			Set objNode = GetItemByValue(value)

			If Not objNode Is Nothing Then
				IsSelectedByValue = objNode.Attributes.getNamedItem("S").nodeValue
			End If

		End Function


		Public Function SetValue(ByVal index , ByVal value )
			If index >= 0 And index < Count Then
				mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("V").nodeValue = value
			End If
		End Function

		Public Function SetText(ByVal index , ByVal value )
			If index >= 0 And index < Count Then
				mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("T").nodeValue = value
			End If
		End Function

		Public Sub SetSelected(ByVal index , ByVal bolSelected )

			If mMode = 0 And bolSelected = True Then
				Me.SetAllSelected False
			End If

			If index >= 0 And index < Count Then
				mobjXMLDocument.firstChild.childNodes(index).Attributes.getNamedItem("S").nodeValue = bolSelected
			End If
		End Sub

		Public Sub SetSelectedByText(ByVal Text , ByVal newvalue )

			Dim objNode

			Set objNode = GetItemByText(Text)

			If mMode = 0 Then
				Me.SetAllSelected False
			End If

			If Not objNode Is Nothing Then
				objNode.Attributes.getNamedItem("S").nodeValue = newvalue
			End If
		End Sub

		Public Sub SetSelectedByValue(ByVal value , ByVal newvalue )
			Dim objNode

			Set objNode = GetItemByValue(value)

			If mMode = 0 Then
				Me.SetAllSelected False
			End If

			If Not objNode Is Nothing Then
				objNode.Attributes.getNamedItem("S").nodeValue = newvalue
			End If

		End Sub

		Private Function GetItemByValue(ByRef value )
			Dim strCritera 
			Dim objNode

			strCritera = "//IC/I[@V='" & value & "']"

			Set GetItemByValue = mobjXMLDocument.selectSingleNode(strCritera)

		End Function

		Private Function GetItemByText(ByRef Text )
			Dim strCritera
			Dim objNode

			strCritera = "//IC/I[@T='" & Text & "']"

			Set GetItemByText = mobjXMLDocument.selectSingleNode(strCritera)

		End Function

		Public Sub SetAllSelected(ByVal bolSelected)

			Dim node
			Dim nodes 'As MSXML2.IXMLDOMNodeList
			Dim strCritera  'As String

			If bolSelected Then
				strCritera = "//IC/I[@S='0']"
			Else
				strCritera = "//IC/I[@S='-1']"
			End If

			Set nodes  = mobjXMLDocument.selectNodes (strCritera)
			For Each node In nodes
				node.Attributes.getNamedItem("S").nodeValue = bolSelected
			Next

			Set node = Nothing
			Set nodes = Nothing

			'For Each node In mobjXMLDocument.firstChild.childNodes
			'	node.Attributes.getNamedItem("S").nodeValue = bolSelected
			'Next

		End Sub

	End Class

%>