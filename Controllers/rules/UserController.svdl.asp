<?xml version="1.0"?>
<ruleset>
	<name>User controller</name>
	<path></path>
	<extraHeaderAction>ignore</extraHeaderAction>
	<extraCookieAction>ignore</extraCookieAction>
	<extraParameterAction>ignore</extraParameterAction>

 	<rule>
		<name>JSESSIONID</name>
		<paramType>cookie</paramType>
		<regex>^[A-F0-9]{32}$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule>

	<rule>
		<name>referer</name>
		<paramType>header</paramType>
		<regex>^http.*$</regex>
		<malformedAction>continue</malformedAction>
		<malformedMessage>Session cookie tampering deteted</malformedMessage>
		<missingAction>continue</missingAction>
	</rule>

 	<rule>
		<name>firstName</name>
		<paramType>PARAMETER</paramType>
		<regex>^[a-zA-Z]{32}$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule>

 	<rule>
		<name>lastName</name>
		<paramType>PARAMETER</paramType>
		<regex>^[a-zA-Z]{32}$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule>

 	<rule>
		<name>token</name>
		<paramType>PARAMETER</paramType>
		<regex>^[a-zA-Z0-9\s.\-]+$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule>


<!--
	<rule>
		<name>password</name>
		<paramType>parameter</paramType>
		<regex>^[0-9]{6}$</regex>
		<malformedAction>continue</malformedAction>
		<malformedMessage>Please correct your password</malformedMessage>
		<missingAction>continue</missingAction>
		<missingMessage>You must enter a password</missingMessage>
		<hidden>true</hidden>
	</rule>
 	-->
</ruleset>