<!-- #include File="_Controller.asp" -->
<!-- #include File="../Models/UserModel.asp" -->

<%
	Dim userController : Set userController = New UserControllerClass

	Call controller.AcceptRequest(userController)

	Class UserControllerClass
		Public Sub ProcessRequest
			Dim lastName : lastName = input.GetString("lastName")
			controller.controllerResponse = "{""lastName"": """ & output.Prepare(lastName) & """}"
		End Sub
	End Class
 %>