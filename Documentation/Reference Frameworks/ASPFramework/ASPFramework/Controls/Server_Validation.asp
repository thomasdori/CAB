<%

	CONST VALIDATION_DATATYPE_STRING = 0
	CONST VALIDATION_DATATYPE_NUMBER = 1
	CONST VALIDATION_DATATYPE_DATE   = 2

	Page.RegisterLibrary "ServerValidationSummary"
	
	Class ServerValidationSummary
		
		Private  mbolIsValid
		Private  mobjMessage
		Public   MessageExplanation
		
		Private Sub Class_Initialize()			
			Set mobjMessage  = New StringBuilder
			mbolIsValid = True
	   End Sub
	   	
	   Private Function GetValidationLink(strObjName, strMessage)
			GetValidationLink = "<LI><A HREF='#' onClick='document.frmForm." & strObjName & ".focus()'>" & strMessage & "</A><BR></LI>"
	   End Function

	   Public Function AddWarning(ErrorMessage)
			mobjMessage.Append "<LI>" & ErrorMessage & "</LI>"
			mbolIsValid = False
	   End Function
	   
	   Public Function AddError(ErrorMessage)
			mobjMessage.Append "<LI><span style='color:red'>" & ErrorMessage & "</span></LI>"
			mbolIsValid = False
	   End Function
	   
	   Public Function Validate(object,strFriendlyName, intDataType, bolRequired, strDefaultValue)
	   			Dim value
				
				'value =  Request.Form(object.Control.Name)
				value = object.Text
				If bolRequired Then
					If value = "" Or value = strDefaultValue Then
						mobjMessage.Append  GetValidationLink(object.Control.Name,strFriendlyName & " is required.")				
						mbolIsValid = False
						Validate = False
						Exit Function
					End If
				End If				
				
				If value = "" Then
					Exit Function
				End If
				
				Select Case intDataType
					Case VALIDATION_DATATYPE_NUMBER:
						If Not IsNumeric(value) Then
							mobjMessage.Append GetValidationLink(object.Control.Name,strFriendlyName & " is not a valid number.")
							mbolIsValid = False
						    Validate = False
						End If	
					Case VALIDATION_DATATYPE_DATE:
						If Not IsDate(value) Then
							mobjMessage.Append GetValidationLink(object.Control.Name,strFriendlyName & " is not a valid date.")
						    mbolIsValid = False
						    Validate = False
						End If	
				End Select		

	   End Function
	   
	   Public Property Get IsValid()
			IsValid = mbolIsValid
	   End Property
	   
	   Public Function GetValidationMessage()
			GetValidationMessage = mobjMessage.ToString()
	   End Function
	   
	   Public Default Function Render()
			'Response.Write "The following validation errors were found. Please correct.<BR>"
			If Not mbolIsValid Then
				Response.Write "<DIV style='width:100%;'><B>Validation Errors:</B>"
				If MessageExplanation<>"" Then
					Response.Write "<BR>"
					Response.Write MessageExplanation
				End If
				Response.Write "<HR>"
				Response.Write mobjMessage.ToString()
				Response.Write "<HR></DIV>"
			End If
	   End Function
	   
	   
	End Class

%>