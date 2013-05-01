<!-- #include File="../config.asp" -->
<!-- #include File="../Helpers/EncodingHelper.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include File="../Helpers/InputHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include File="../Helpers/ValidationHelper.asp" -->
<!-- #include File="../Models/_Model.asp" -->

<%
	Dim controller : Set controller = New ControllerClass

	Class ControllerClass
		Public controllerResponse

		Private Sub Class_Initialize()
 			controllerResponse = ""
		End Sub

		Public Sub AcceptRequest(childController)
			If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			    'Prevent CSRF (Cross-Site Request Forgeries) by comparing request-generated tokens.
			    If Request.Form(csrf.GetTokenKey()) = Session(csrf.GetTokenKey()) Then
			    	' Create new token
					csrf.SetToken()
			    Else
			    	error.Handle(CSRF_ERROR_CODE)
			    End If

			    'Check Request Prameter with Stinger
 				Dim errors : Set errors = validator.Validate()

				If Not errors Is nothing And errors.count > 0 Then
					'todo: return validator.format(errors)
		      	End if
			Else
				error.Handle(UNALLOWED_HTTP_METHOD_ERROR_CODE)
			End If

 			Call childController.ProcessRequest()

 			WriteResponse()
		End Sub

		Private Sub WriteResponse()
			Response.ContentType = "application/json"
			Response.Write("{""token"": """ & csrf.GetToken() & """, ""response"": " & controllerResponse & "}")
		End Sub
	End Class
%>