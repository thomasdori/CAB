<!-- #include File="../config.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/InputHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<%
	' todo: csrf + validation'

	Dim controller
	Set controller = New ControllerClass

	Class ControllerClass
		Public Sub AcceptRequest(childController)
			'Send Javascript Code for better user experience, pass form name as parameter
 			output.Write(validator.getJavaScript ("form"))

 			Call childController.ProcessRequest
		End Sub
	End Class
%>