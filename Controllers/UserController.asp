<!-- #include File="Controller.asp" -->
<!-- #include File="../Models/UserModel.asp" -->

<%
	Dim userController : Set userController = New UserControllerClass

	Call controller.AcceptRequest(userController)

	Class UserControllerClass
		Public Sub ProcessRequest
			Dim name
			name = Request("firstName") 'input.GetParameter("firstName")

			output.WriteLine("firstName: " & firstName)
		End Sub
	End Class
 %>