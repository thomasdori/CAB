<!-- #include File="../Models/ErrorModel.asp" -->

<%
	'Dependencies: OutputHelper.asp
	Const CSRF_ERROR_CODE = "CSRF_ERROR_CODE"
	Const UNALLOWED_HTTP_METHOD_ERROR_CODE = "UNALLOWED_HTTP_METHOD_ERROR_CODE"

	Dim error : Set error = New ErrorHelperClass

	Class ErrorHelperClass
		Public Function Handle(errorCode)
			' Check if the request is an ajax request
			If (Request.ServerVariables("HTTP_X-Requested-With") = "XMLHttpRequest") Then
				Select Case errorCode
					Case CSRF_ERROR_CODE, UNALLOWED_HTTP_METHOD_ERROR_CODE
						Response.Redirect("/Views/Errors/InvalidRequestError.asp")
					Case Else
						Response.Redirect("/Views/Errors/UnknownError.asp")
				End Select
			Else
				' TODO: generate error output for ajax requests
				output.WriteLine("An error occured.")
			End If
		End Function

		Public Function ThrowError
			'Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(header) + "] header was unexpected"
		End Function

		Public Function getErrors
			'clear buffer'
			'write error message
			'response flush'
			GetErrors = ""
		End Function

		Public Function GetCustomErrors
			'Custom Error Handling
			If (Session("error") <> "") Then
				'todo: display a pupup or inline error message
				GetCustomErrors = ""

				'clear the error variable
				Session("error") = ""
			End If
		End Function

		Public Function GetAspErrors
			'ASP Error Handling
			Dim oError
			Set oError = Server.GetLastError()

			If (oError.Number <> 0) Then
				errorModel.write(oError)
			    oError.Clear
			End If

			GetAspErrors = ""
		End Function
	End Class
%>