<!-- #include File="../Helpers/ActionHelper.asp" -->

<%
	Dim controller
	Set controller = New ControllerClass

	Class ControllerClass
		Public Sub AcceptRequest
			action.Execute("ProcessRequest")
		End Sub
	End Class
%>