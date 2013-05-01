<%

' ---------------------------------------------------------------
'  Error Helpers
' ---------------------------------------------------------------
' 
' send_error_email
' build_error_email


' ---------------------------------------------------------------
'  Send Error Emails
' ---------------------------------------------------------------
' 
' @var = string

function send_error_email(strFile, err_number, err_source, err_description, err_helpContext)
	
	dim local_str, LocalText, objCDOMail, count, Item, createTable
	dim strBody, x, y, MachineName,  pc
	
	Set pc = CreateObject("Wscript.Network") 
    
		MachineName = pc.ComputerName
    
	Set pc = nothing 
	
	local_str = ""&Vbcrlf
	
	local_str = local_str & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"&Vbcrlf
	
	local_str = local_str & "<html>"&Vbcrlf
	
	local_str = local_str & "<head>"&Vbcrlf
	
	local_str = local_str & "<style>"&Vbcrlf
	
	local_str = local_str & "body, td, th, p, li { font : 12px Arial, Helvetica, sans-serif; }"&Vbcrlf
	
	local_str = local_str & "h2 { font-weight: bold; padding: 0; margin: 0;font : 16px Arial, Helvetica, sans-serif; }"&Vbcrlf
	
	local_str = local_str & "a { color: #990000;}"&Vbcrlf
	
	local_str = local_str & "#dataTable {width: 600px;}"&Vbcrlf
	
	local_str = local_str & "#dataTable caption {font-style: italic; font-size: 10px;  color: #999; text-align: right;padding:5px;}"&Vbcrlf
	
	local_str = local_str & "#dataTable th {font-weight:bold; border-top: 1px solid #CCC;border-bottom: 1px solid #CCC;border-right: 1px solid #CCC;text-align: center;padding: 5px 8px;background: #5C7C9C;color:#fff;}"&Vbcrlf
	
	local_str = local_str & "#dataTable th.nobg {border-top: 1px solid #CCC;border-right: 1px solid #CCC;border-left: 1px solid #CCC;border-bottom: 1px solid #CCC;background: #5C7C9C;color:#fff;}"&Vbcrlf
	
	local_str = local_str & "#dataTable td {border-right: 1px solid #CCC;border-bottom: 1px solid #CCC;background: #fff;padding: 5px 8px;}"&Vbcrlf
	
	local_str = local_str & "#dataTable td.alt {background: #f7f7f7;}"&Vbcrlf
	
	local_str = local_str & "#dataTable td.spec {font-weight:normal; color:#333;text-align: left;border-top: none;border-left: 1px solid #CCC;border-right: 1px solid #CCC;border-bottom: 1px solid #CCC;background: #fff;}"&Vbcrlf
	
	local_str = local_str & "#dataTable td.specalt {font-weight:normal; text-align: left;border-top: none;border-left: 1px solid #CCC;border-right: 1px solid #CCC;border-bottom: 1px solid #CCC;background: #f7f7f7; color:#333;}"&Vbcrlf
	
	local_str = local_str & "</style>"&Vbcrlf
	
	local_str = local_str & "</head>"&Vbcrlf
	
	local_str = local_str & "<body>"&Vbcrlf
	
	local_str = local_str & "<h2>Errors on: " & config_error_name & "</h2>"&Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "Error Reported on " &  formater(now, "General Date") & Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "On Page: http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("HTTP_SCRIPT_URL")& Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "FATAL ERROR: while trying to execute: <b>" & strFile & "</b>" &Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str &  "Error description: " & err_description &Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "Error Source: " & err_source&Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "Server Name: " & MachineName&Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	local_str = local_str & "<br />"&Vbcrlf
	
	
	'--------------------------------------------------------------
	'	Display Request.Form Table in EMail.
	'--------------------------------------------------------------
	If Request.Form <> "" Then
		
		Dim value4() 
    	
		ReDim value4(Request.Form.Count - 1,1)
		
		i = 0
		
		for each Item In Request.Form
    		
			value4(i,0) =  Item
			
			value4(i,1) = Request.Form(Item)
			
			i = i + 1
			
		next
		
		Set createTable = New htmlTable
		
			createTable.headers 	= array("Form Name","Form Value")
			
			createTable.altRows 	= true
			
			createTable.array_data 	= value4
			
			local_str = local_str & createTable.build()
			
		Set createTable = nothing
		
		local_str = local_str & "<br />"&Vbcrlf
	
	end if
	
	
	
	'--------------------------------------------------------------
	'	Display Cookies Table in EMail.
	'--------------------------------------------------------------
	
	If Request.Cookies <> "" Then
		
		Dim value1() 
    	
		ReDim value1(Request.Cookies.Count - 1,1)
		
		i = 0
		
		for each x in Request.Cookies
			
			if Request.Cookies(x).HasKeys then
				
				for each y in Request.Cookies(x)
					
					value1(i,0) =  x & ":" & y
					value1(i,1) = Replace(Request.Cookies(x)(y),"%","% ")
					
				next
				
			else
				
				value1(i,0) =  x
				value1(i,1) =  Replace(Request.Cookies(x),"%","% ")
				
			end if
			
			i = i + 1
			
		next
			
			Set createTable = New htmlTable
				
				createTable.headers 	= array("Cookie Name","Cookie Value")
				
				createTable.array_data 	= value1
				
				local_str = local_str & createTable.build()
				
			Set createTable = nothing
		
		local_str = local_str & "<br />"&Vbcrlf
		
	end if
	
	
	
	'--------------------------------------------------------------
	'	Display Session Table in EMail.
	'--------------------------------------------------------------
	
	If Session.Contents.Count <> 0 Then
	
		count=Session.Contents.Count - 1
	
		Dim value3() 
    	
		ReDim value3(count,1)
		
		i = 0
		
		for each Item in Session.Contents
			value3(i,0) =  Item
			
			value3(i,1) = Session.Contents(Item)
			
			i = i + 1
			
		next
		
		Set createTable = New htmlTable
		
			createTable.headers 	= array("Session Name","Session Value")
			
			createTable.array_data 	= value3
			
			local_str = local_str & createTable.build()
			
		Set createTable = nothing
		
		local_str = local_str & "<br />"&Vbcrlf
		
	end if
	
	
	count = Request.ServerVariables.Count - 1
		
	Dim value2() 
    
	ReDim value2(count,1)
	
	i = 0
	
	For Each Item in Request.ServerVariables
   		
		value2(i,0) =  Item
		
		value2(i,1) = Replace(Request.ServerVariables(Item),"=","= ")
	
		i = i + 1
	
	Next

	Set createTable = New htmlTable
	
		createTable.headers 	= array("Server Varible Name","Server Varible Value")
		
		createTable.altRows 	= true
		
		createTable.array_data 	= value2
		
		local_str = local_str & createTable.build()
		
	Set createTable = nothing
	
	local_str = local_str & "<p>Powered by: <em><strong>RIS Core</strong></em></p>"&Vbcrlf
	
	local_str = local_str & ""&Vbcrlf
	
	local_str = local_str & "</body>"&Vbcrlf
	
	local_str = local_str & "</html>"&Vbcrlf
	
	
	
	
	'--------------------------------------------------------------
	'	Create Email for Text
	'--------------------------------------------------------------
	
	LocalText = LocalText & "<h2>Errors on: " & config_error_name & "</h2>"&Vbcrlf
	
	LocalText = LocalText & "<br />"&Vbcrlf
	
	LocalText = LocalText & "Error Reported on " &  formater(now, "General Date") & Vbcrlf
	
	LocalText = LocalText & "<br />"&Vbcrlf
	
	LocalText = LocalText & "On Page: http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("HTTP_SCRIPT_URL") & Vbcrlf
	
	LocalText = LocalText & "<br />"&Vbcrlf
	
	LocalText = LocalText & "FATAL ERROR: while trying to execute: <b>" & strFile & "</b>" &Vbcrlf
	
	LocalText = LocalText & "<br />"&Vbcrlf
	
	LocalText = LocalText &  "Error description: " & err_description &Vbcrlf
	
	LocalText = LocalText & "<br />"&Vbcrlf
	
	LocalText = LocalText & "Error Source: " & err_source&Vbcrlf
	
	LocalText = remove_html(LocalText)
	
	
	
	
	if EMAILERRORS then
		
		if robot_check = false then
			
			Call email_send(array("to",config_error_email) , config_error_name & " Errors Detected", LocalText, local_str, "")
		
		end if
		
	end if

end function

%>