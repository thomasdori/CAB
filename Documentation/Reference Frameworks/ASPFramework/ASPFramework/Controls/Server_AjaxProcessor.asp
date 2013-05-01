<%Option Explicit%>
<!--#Include File = "StringBuilder.asp"-->
<!--#Include File = "Server_BrowserAgent.asp"-->
<!--#Include File = "CLASP_Setup.asp"-->
<%
	Dim AjaxForm 
	Set AjaxForm = new cAjax
	AjaxForm.ProcessRequest
	
	Class cAjax
		Private AjaxXml
		Dim Action
		Dim ActionSource
		Dim ActionSourceInstance
		Dim ActionXMessage				

		Private Sub Class_Initialize()		

		End Sub
		
		Public Function ProcessRequest()
				If Request.Form("AjaxPost") = "1" Then				

					Action       = Request.Form("AjaxAction")
					ActionSource = Request.Form("AjaxSource")
					
					If Request.Form("AjaxMode") = "1" Then
						Response.Clear
						ExecuteEventFunctionParams Action,ActionSource
						Response.End
					End If
					Set AjaxXml  = CreateObject(CLASP_DOM_PARSER_CLASS)			
					If Not AjaxXml.LoadXml(Request.Form("AjaxForm")) Then
						Response.Clear
						Response.Write "Error Processing the AJAX Form" & Server.HTMLEncode(Request.Form("AjaxForm"))
						Response.End
					End If
					
					Response.Clear
					ExecuteEventFunctionParams Action,ActionSource
					Response.End
				
			End If

			Response.Clear
			Response.Write "Invalid Ajax Post."
			Response.End	
	
		End Function
		
		Public Function GetFormValue(FieldName)
			Dim objFields
			Set objFields = AjaxXml.selectNodes("//" & FieldName)
			
			If objFields.length > 0	Then	
				GetFormValue =   URLDecode(objFields.item(0).text)
			End If

			Set objFields = Nothing

		End Function
		
		Public Function ExecuteEventFunctionParams(EventName,p)
			Dim fnc		
			ExecuteEventFunctionParams = False
			Set fnc = GetFunctionReference(EventName)		
			If Not fnc Is Nothing Then
				ExecuteEventFunctionParams = True			
				Call fnc(p)
			End If
		End Function	

		Private Function GetFunctionReference(sFncName)
			On Error Resume Next
				Set GetFunctionReference = Nothing
				Set GetFunctionReference = GetRef(sFncName)		
			On error Goto 0
		End Function

		Public Function URLDecode(str)
			dim re
			set re = new RegExp

			str = Replace(str, "+", " ")
			
			re.Pattern = "%([0-9a-fA-F]{2})"
			re.Global = True
			URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
		End Function

	End Class


	' Replacement function for the above
	Function URLDecodeHex(match, hex_digits, pos, source)
		URLDecodeHex = chr("&H" & hex_digits)
	End function

%>

