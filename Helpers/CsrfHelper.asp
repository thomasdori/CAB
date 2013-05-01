<!-- #include File="Md5Helper.asp" -->
<!-- #include File="GuidHelper.asp" -->

<%
	'Dependencies: ErrorHandler.asp
	Const TOKEN_KEY = "token"

	Dim csrf : Set csrf = New CsrfHelperClass

	Class CsrfHelperClass
		Public Function GetToken()
			GetToken = Session(TOKEN_KEY)
		End Function

		Public Sub SetToken()
			Session(TOKEN_KEY) = NewToken()
		End Sub

		' According to the Synchronizer Token Pattern
		' Source: http://stackoverflow.com/questions/6421417/howto-implement-synchronizer-token-pattern-in-classic-asp
		Public Function DoesTokenExist
			If(GetToken() = "") Then
				DoesTokenExist = False
			End If

			DoesTokenExist = True
		End Function

		Public Function GetTokenKey()
			GetTokenKey = TOKEN_KEY
		End Function

		Private Function NewToken()
			NewToken = md5(guid.GetGuid())
		End Function
	End Class
%>