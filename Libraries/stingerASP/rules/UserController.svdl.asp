<?xml version="1.0"?>
<ruleset>
	<name>Login Form</name>
	<path>/Stinger-1.0r1/test</path>
	<extraHeaderAction>ignore</extraHeaderAction>
	<extraCookieAction>continue</extraCookieAction>
	<extraParameterAction>ignore</extraParameterAction>

 	<rule>
		<name>JSESSIONID</name>
		<paramType>cookie</paramType>
		<regex>^[A-F0-9]{32}$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>ignore</missingAction>
	</rule>

<!--
	<rule>
		<name>referer</name>
		<paramType>header</paramType>
		<regex>^http.*$</regex>
		<malformedAction>continue</malformedAction>
		<malformedMessage>Session cookie tampering deteted</malformedMessage>
		<missingAction>ignore</missingAction>
	</rule>

	<rule>
		<name>username</name>
		<paramType>parameter</paramType>
		<regex>^[\w]{6}$</regex>
		<malformedAction>continue</malformedAction>
		<malformedMessage>Please correct your username</malformedMessage>
		<missingMessage>You must enter a username</missingMessage>
		<missingAction>continue</missingAction>
	</rule>

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

	<rule>
		<name>domain</name>
		<paramType>parameter</paramType>
		<regex>^$|[\w]{4,8}$</regex>
		<malformedAction>continue</malformedAction>
		<messageLevel>detailed</messageLevel>
		<missingAction>ignore</missingAction>
	</rule>
 	-->
</ruleset>