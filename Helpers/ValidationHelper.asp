<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<%
	'http://stackoverflow.com/questions/15723355/stingerasp-cant-seem-to-validate-referer-header-rule-specified-in-svdl-file'
	'Basically, the continue parameter for any action should stop further processing of the ASP page and halt with an appropriate error.
	'The ignore parameter would ignore the rule and proceed with further processing.
	'paramType could either be header, cookie or parameter (GET or POST, or I guess both)
	'regex matches the contents of the parameter to be tested, against a regular expression

	Dim stingerObj 	: Set stingerObj = New Stinger
	Dim validator 	: Set validator = New ValidationHelperClass

	Class ValidationHelperClass
		Public Function Validate()
			Validate = stingerObj.Validate()
		End Function
	End Class
%>