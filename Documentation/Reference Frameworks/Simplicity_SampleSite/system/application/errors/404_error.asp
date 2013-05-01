<%

' ---------------------------------------------------------------
'  Config changes for Controller Page
' ---------------------------------------------------------------
'
'

site_name 			= "404 Page Not Found"
default_template 	= "error.asp"

' ---------------------------------------------------------------
' Content for Controller Page
' ---------------------------------------------------------------
' 
' @write = string

change_header("404")

write = write & "<h1><span class=""error"">404 Page Not Found</span></h1>"
write = write & "<p>The page you requested was not found.</p>"

%>
