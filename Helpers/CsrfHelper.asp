<!-- #include File="Md5Helper.asp" -->
<!-- #include File="GuidHelper.asp" -->

<%
	'Dependencies: ErrorHandler.asp
	Const TOKEN_KEY = "token"

	Dim csrf : Set csrf = New CsrfHelperClass

	' According to the Synchronizer Token Pattern
	' Source: http://stackoverflow.com/questions/6421417/howto-implement-synchronizer-token-pattern-in-classic-asp
	Class CsrfHelperClass
		Public Function GetToken()
			If(Session(TOKEN_KEY) = "") Then
				Session(TOKEN_KEY) = NewToken()
			End If

			GetToken = Session(TOKEN_KEY)
		End Function

		Public Sub SetToken()
			Session(TOKEN_KEY) = NewToken()
		End Sub

		Public Function GetTokenKey()
			GetTokenKey = TOKEN_KEY
		End Function

		Private Function NewToken()
			NewToken = md5(guid.GetGuid())
		End Function
	End Class
%>