<!-- #include File="Md5Helper.asp" -->
<!-- #include File="GuidHelper.asp" -->

<%
	'Dependencies: ErrorHandler.asp

	Dim csrf : Set csrf = New CsrfHelperClass

	' Calling the procedure to check if the
	Call csrf.CheckParameter


	Class CsrfHelperClass
		Public Function GetParameter()
			GetParameter = Session("token")
		End Function

		' According to the Synchronizer Token Pattern
		' Source: http://stackoverflow.com/questions/6421417/howto-implement-synchronizer-token-pattern-in-classic-asp
		Public Sub CheckParameter
			' checks to make sure the request method is truly a post-back
			If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			    ' Prevent CSRF (Cross-Site Request Forgeries) by comparing request-generated tokens.
			    If Request.Form("token") = Session("token") Then
			    	' Create new token
					Session("token") = md5(guid.GetGuid())
			    Else
			    	error.Handle(CSRF_ERROR_CODE)
			    End If
			End If
		End Sub
	End Class
%>