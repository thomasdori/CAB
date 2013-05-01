<%

' ---------------------------------------------------------------
'  Config changes for Controller Page
' ---------------------------------------------------------------
'
'

site_name 			= "Application Error"
default_template 	= "error.asp"

' ---------------------------------------------------------------
' Content for Controller Page
' ---------------------------------------------------------------
' 
' @write = string

write = write & "<h1>ASP Application Scripting Error</h1>"  & error_title
write = write & "<p>FATAL ERROR: while trying to execute: <strong><i>" & strFile & "</i></strong></p>" 

%>