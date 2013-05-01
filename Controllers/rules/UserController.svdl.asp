<?xml version="1.0"?>
<ruleset>
	<name>User controller</name>
	<path></path>
	<extraHeaderAction>ignore</extraHeaderAction>
	<extraCookieAction>ignore</extraCookieAction>
	<extraParameterAction>ignore</extraParameterAction>

 	<rule>
		<name>firstName</name>
		<paramType>PARAMETER</paramType>
		<regex>^[a-zA-Z0-9\s.\-]+$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule>

 	<rule>
		<name>lastName</name>
		<paramType>PARAMETER</paramType>
		<regex>^[a-zA-Z0-9\s.\-]+$</regex>
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

<!-- For Views -->
<!--
	<rule>
		<name>referer</name>
		<paramType>header</paramType>
		<regex>^http.*$</regex>
		<malformedAction>continue</malformedAction>
		<malformedMessage>Session cookie tampering deteted</malformedMessage>
		<missingAction>continue</missingAction>
	</rule>
 -->

<!-- Not working since asp session id key contains a random string at the end -->
<!--   	<rule>
		<name>ASPSESSIONID</name>
		<paramType>cookie</paramType>
		<regex>^[A-F0-9]{32}$</regex>
		<malformedAction>continue</malformedAction>
		<missingAction>continue</missingAction>
	</rule> -->
</ruleset>