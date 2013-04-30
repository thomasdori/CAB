<!-- #include File="../Helpers/ConfigurationHelper.asp" -->
<!-- #include File="../Helpers/InputHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<!-- #include File="../Models/UserModel.asp" -->

<%
	Dim userController
	Set userController = new UserController
	Call userController.ProcessRequest

	Class UserController

		Sub ProcessRequest
			output.write("test0")
			output.write(Request.ServerVariables("HTTP_X-Requested-With"))
			output.write("test1")
			'Send Javascript Code for better user experience, pass form name as parameter
 			output.write(validator.getJavaScript ("form"))
 			output.write("test1")

			'Dim name = input.getParameter("name")
		End Sub
	End Class
 %>