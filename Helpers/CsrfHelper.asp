<!-- #include File="Helpers/Md5Helper.asp" -->

<%
	Dim csrf
	Set csrf = new CsrfHelper

	' Calling the procedure to check if the
	Call csrf.checkParameter


	Class CsrfHelper
		Function getParameter()
			getParameter = Session("token")
		End Function

		' According to the Synchronizer Token Pattern
		' Source: http://stackoverflow.com/questions/6421417/howto-implement-synchronizer-token-pattern-in-classic-asp
		Sub checkParameter
			' checks to make sure the request method is truly a post-back
			If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			    ' Prevent CSRF (Cross-Site Request Forgeries) by comparing request-generated tokens.
			    If Request.Form("token") = Session("token") Then
			    	' Create new token
					Session("token") = md5(GetGUID())
			    Else
			    	error.handle(CSRF_ERROR_CODE)
			    End If
			End If
		End Sub
	End Class
%>