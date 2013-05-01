<!--#Include File = "Server_CoreObjects_UseASP.asp"-->
<%
	Response.buffer=True
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.ExpiresAbsolute = Now()- 1000
	Response.Expires = 0
CONST WEB_CONTROLS_PATH	    = "/ASPFramework/Controls/"
CONST SCRIPT_LIBRARY_PATH	= "/ASPFramework/Controls/ScriptLibrary/"
'CONST CLASP_DOM_PARSER_CLASS = "MSXML2.FreeThreadedDOMDocument.2.6" '(or 3.0)
CONST CLASP_DOM_PARSER_CLASS = "MSXML2.DOMDocument.2.6" '(or 3.0)
'CONST CLASP_DOM_PARSER_CLASS = "MSXML.DOMDocument" '(or 3.0)
%>