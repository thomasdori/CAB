<!-- #include File="../Helpers/EventHelper.asp" -->

<%
	Class Page
		Public AuthorizedRoles
		Public RequiresAuthentication
		Public RequiresAuthorization
		Public PageTitle
		Public ScriptFile
		Public StyleFile
		Public ContentPartial

		Private Sub Class_Initialize
			'set scriptFile to Request.ServerVariables("SCRIPT_NAME") & ".js"
			'set styleFile to Request.ServerVariables("SCRIPT_NAME") & ".css"
			'set contentpartial to Request.ServerVariables("SCRIPT_NAME") & ".html"
		End Sub

		Public Function OpenPage()
			Call  eventHelper.Execute("Page_Configure")
		End Function

		Public Function ClosePage()

		End Function

		Public Function RenderPage()
			Call OpenPage()
			Call GetContent()
			Call ClosePage()
		End Function

	End Class
%>