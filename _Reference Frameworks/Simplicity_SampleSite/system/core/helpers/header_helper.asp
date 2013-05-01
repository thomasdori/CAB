<%

' ---------------------------------------------------------------
'  Cookie Helper
' ---------------------------------------------------------------
' 
' change_header


' ---------------------------------------------------------------
'  change_header
' ---------------------------------------------------------------
' 
'  @string

function change_header(var)
	
	select case lcase( var )
	
		case "xml"
			Response.Status = "200 OK"
			Response.ContentType = "text/xml"
		
		case "404"
			Response.Status = "404 Not Found"
	end select
end function

%>