<%

' ---------------------------------------------------------------
'  Redirect Helpers
' ---------------------------------------------------------------
' 
' redirect

' Legend
'	200 - OK
'	301 - Moved Permanently
'	302 - Found
'	404 - Not Found

' ---------------------------------------------------------------
'  Send Error Emails
' ---------------------------------------------------------------
' 
' @header_url = string
' @header_type = integer

Sub redirect(header_url, header_type)
	
	dim page_text, page_title, check_html
	
	Response.Buffer = True
	
	Select Case header_type
		case 200
    		check_html	= false
			Response.Status = "200 OK"
			Response.redirect(header_url)
		
		case 301
			check_html	= true
			page_title 	= "301 Moved Permanently"
			page_text 	= "The page you want has been moved"
			Response.Status = "301 Moved Permanently"
		
		case 302
   			check_html	= true
			page_title 	= "302 Found"
			page_text 	= "The page you requested is currently unavailable"
			Response.Status = "302 Found"
		
		case 404
			check_html	= true
			page_title 	= "404 Not Found"
			page_text 	= "Sorry, that page does not exists"
			Response.Status = "404 Not Found"
		
		case else
			check_html	= false
			Response.Status = "200 OK"
			Response.redirect(header_url)
		
	end select
	
	if check_html	= true then
		Response.AddHeader "Location", header_url
		Response.Clear()
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ""http://www.w3.org/TR/html4/strict.dtd"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>Redirection</title>"
		Response.Write "<META HTTP-EQUIV=Refresh CONTENT=""0;URL=" & header_url & """>"
		Response.Write "<Script>window.location='" & header_url & "';</Script>"
		Response.Write "</head>"
		Response.Write "<body>"
		Response.Write "<h1>" & page_title & "</h1>"
		Response.Write "<p>" & page_text & "</p>"
		Response.Write "</body>"
		Response.Write "</html>"
	end if
	
end Sub

%>