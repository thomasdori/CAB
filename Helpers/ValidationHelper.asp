<!-- #include file="../Libraries/StingerASP/stinger.asp" -->

<%
	'Help for defining a svdl file:'
		'The main rule:
			'name
			'path: ???
			'extraHeaderAction: what to do if another parameter was sent
			'extraCookieAction: what to do if another parameter was sent
			'extraParameterAction: what to do if another parameter was sent
		'The rules'
			'name: The parameter name
			'paramType: HEADER (request.ServerVariables), COOKIE (request.Cookies), PARAMETER (Request.QueryString and Request.Form)
			'regex: The regular expression used to validate
			'extraHeaderAction: FATAL [default] (causes err.raise), CONTINUE (adds error to array which will be returned), IGNORE (ignores the error)
			'extraHeaderMessage: Arbitrary message
			'missingAction: FATAL (default), CONTINUE, IGNORE
			'missingMessage: Arbitrary message
			'hidden: ???

		'Some further explaination: http://stackoverflow.com/questions/15723355/stingerasp-cant-seem-to-validate-referer-header-rule-specified-in-svdl-file

	'paramType could either be header, cookie or parameter (GET or POST, or I guess both)
	'regex matches the contents of the parameter to be tested, against a regular expression

	Dim stingerObj 	: Set stingerObj = New Stinger
	Dim validator 	: Set validator  = New ValidationHelperClass

	Class ValidationHelperClass
		Public Function Validate()
			'Set is neccesary here
			Set Validate = stingerObj.Validate()
		End Function
	End Class
%>