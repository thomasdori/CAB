<!-- #include File="../Helpers/ConfigurationHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<!-- #include File="../Models/UserModel.asp" -->

<%
	Class UserController

		Sub ProcessRequest
		End Sub

	End Class

 	'Send Javascript Code for better user experience, pass form name as parameter
 	output.write(validator.getJavaScript ("form"))

	'Dim name = input.getParameter("name")
 %>