<%

' ---------------------------------------------------------------
'  Config changes for Controller Page
' ---------------------------------------------------------------
'
'

site_name 			= "URL Error"
default_template 	= "error.asp"

' ---------------------------------------------------------------
' Content for Controller Page
' ---------------------------------------------------------------
' 
' @write = string

change_header("404")

write = write & "<h1><span class=""error"">URL Website Error</span></h1>"
write = write & "<p>The URL you submitted has disallowed characters.  Please check your URL, and try again.</p>"

%>
