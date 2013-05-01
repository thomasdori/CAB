<!--#Include File = "Server_ASPCrypto3.asp" -->
<%
Class cASPViewState
	Private mobjXMLDocument
	Private mobjNode
	Private mCurrentNode


	Public Sub GetDomObject(ByRef obj)
	    
		Set obj = mobjXMLDocument
	    
	End Sub


	Private Sub Class_Initialize()
		Call InitXML
	End Sub

	Private Sub InitXML()		
		'If need to support special characters...
		'http://support.microsoft.com/default.aspx?scid=kb;EN-US;315580
		'Set mobjXMLDocument = CreateObject("MSXML2.FreeThreadedDOMDocument.3.0")
		Set mobjXMLDocument = CreateObject(CLASP_DOM_PARSER_CLASS)
		
		mobjXMLDocument.loadXML "<VS><V N=""""/></VS>"
	    Set mobjNode = mobjXMLDocument.firstChild.childNodes(0).cloneNode(False)
		mobjXMLDocument.firstChild.removeChild mobjXMLDocument.firstChild.childNodes(0)		
	End Sub

	Public Sub Clear()
		Call InitXML
	End Sub

	Public Sub LoadViewState(ByRef strViewStateXML)	
		If Not mobjXMLDocument.loadXML(strViewStateXML) Then
			Err.Raise 50000,"LoadViewState","Error Loading ViewState"
		End If
	End Sub
	
	Public Sub LoadViewStateBase64(ByRef strViewStateXMLBase64,ByVal bolIsCompressed)
	            
			Dim objCrypto
			Dim strXML		
	        
			strXML = strViewStateXMLBase64
			If Len(strXML) = 0 Then
				Call InitXML
			Else
				Set objCrypto = New cASPCrypto									
				strXML = objCrypto.DecodeStr64(strXML)	            	        
	            If Not mobjXMLDocument.loadXML(strXML ) Then					
					Call InitXML
				End If			
				Set objCrypto = Nothing

			End If
	        
			'strViewStateXMLBase64 = strXML
			
	End Sub
	
	Public Function GetViewState()
			GetViewState = mobjXMLDocument.xml
	End Function
	
	Public Function GetViewStateBase64(lngCompressIfLargerThan, ByRef bolIsCompressed)
			Dim strXML
			Dim objCrypto 
			
			Set objCrypto = New cASPCrypto
	        
			strXML = mobjXMLDocument.xml
	                	        
			GetViewStateBase64 = objCrypto.EncodeStr64(strXML)
	        
			Set objCrypto = Nothing

	End Function

	Public Sub Add(ByVal Name, ByVal value)
		Dim xmlNode
		Set xmlNode = GetNodeByName(Name)	    
		If xmlNode Is Nothing Then
			Set xmlNode = mobjXMLDocument.firstChild.appendChild(mobjNode.cloneNode(False))						  
		End If
	    
		With xmlNode
			.Text = value & ""
			.Attributes.getNamedItem("N").nodeValue = Name
		End With	    
	End Sub
	
	Public Property Get Count()
		Count = mobjXMLDocument.firstChild.childNodes.length
	End Property


	Public Sub Remove(ByVal index)
		Call mobjXMLDocument.firstChild.removeChild(mobjXMLDocument.firstChild.childNodes(index))
	End Sub

	Public Sub RemoveByName(ByVal Name)
		Dim objNode
		Set objNode = GetNodeByName(Name)
		If Not objNode Is Nothing Then
			Call mobjXMLDocument.firstChild.removeChild(objNode)
		End If
	End Sub
	Public Function GetValueByIndex(ByVal index)
		GetValueByIndex = mobjXMLDocument.firstChild.childNodes(index).Text
	End Function

	Public Sub SetValueByIndex(ByVal index, ByVal value)
		mobjXMLDocument.firstChild.childNodes(index).Text = value
	End Sub
	
	Public Function GetValue(ByVal Name)
		Dim objNode
		Set objNode = GetNodeByName(Name)
		If Not objNode Is Nothing Then
			GetValue = objNode.Text
		End If
	End Function

	Public Sub SetValue(ByVal Name, ByVal value)
		Dim objNode
		Set objNode = GetNodeByName(Name)
		If Not objNode Is Nothing Then
			objNode.Text = value
		End If
	        
	End Sub

	Private Function GetNodeByName(ByVal Name)
		Dim strCritera
		Dim objNode
		strCritera = "//VS/V[@N = '" & Name & "']"
		Set GetNodeByName = mobjXMLDocument.selectSingleNode(strCritera)
	End Function

	Public Function HasKey(ByVal Name) 
			HasKey = (Not GetNodeByName(Name) Is Nothing)
	End Function
End Class

%>