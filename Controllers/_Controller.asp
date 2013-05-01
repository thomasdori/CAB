<!-- #include File="../config.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include File="../Helpers/InputHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<%
	' todo: csrf + validation'

	Dim controller : Set controller = New ControllerClass

	Class ControllerClass
		Public Sub AcceptRequest(childController)
			If(Not csrf.DoesTokenExist()) Then
				csrf.CreateToken()
			Else
				If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
				    ' Prevent CSRF (Cross-Site Request Forgeries) by comparing request-generated tokens.
				    If Request.Form(csrf.GetTokenKey()) = Session(csrf.GetTokenKey()) Then
				    	' Create new token
						csrf.SetToken()
				    Else
				    	error.Handle(CSRF_ERROR_CODE)
				    End If
				End If
			End If

			'Send Javascript Code for better user experience, pass form name as parameter
 			output.Write(validator.getJavaScript ("form"))

 			Call childController.ProcessRequest
		End Sub
	End Class
%>