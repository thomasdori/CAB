<%

' ---------------------------------------------------------------
'  Error Helpers
' ---------------------------------------------------------------
' 
' email_send



sub email_send(send_values, subject, text_body, html_body, attachment)
	
	dim objCDOMail, i
	
	Set objCDOMail = Server.CreateObject("CDO.Message")
		
		for i = 0 to ubound(send_values)
			
			if send_values(i) = "cc" then
				
				objCDOMail.Cc      	= send_values(i+1)
				
			elseif send_values(i) = "bcc" then
				
				objCDOMail.Bcc      	= send_values(i+1)
				
			elseif send_values(i) = "from" then
				
				objCDOMail.From    		= send_values(i+1)
				
			else
				
				objCDOMail.To      	= send_values(i+1)
				
			end if
			
			i = i + 1
		
		next
		
		
		
		objCDOMail.Subject 	= subject
		
		objCDOMail.TextBody = text_body
		
		if does_exists(html_body) then
		
			objCDOMail.HTMLBody = html_body
		
		end if
		
		if does_exists(attachment) then
		
			objCDOMail.AddAttachment attachment
		
		end if
		
		objCDOMail.send
		
	Set objCDOMail = Nothing
	
end sub

%>