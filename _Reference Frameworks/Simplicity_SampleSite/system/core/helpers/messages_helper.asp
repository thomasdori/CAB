<%

' ---------------------------------------------------------------
'  Messages Helper
' ---------------------------------------------------------------
' 
' show_flash_messages
' show_static_messages
' show_form_errors


' ---------------------------------------------------------------
'  Show Flash Messages
' ---------------------------------------------------------------
' 
'  @string

function show_flash_messages
	
	dim pItem, p_Item, flash_type, plocal
	
	flash_type = "NO_FLASH"
	
	for each pItem in Request.Cookies
		
		if Request.Cookies(pItem).HasKeys then
		
			for each p_Item in Request.Cookies(pItem)
				
				if instr(p_Item,"_flash") <> 0 then
					
					plocal = plocal & message_box( replace(p_Item, "_flash", ""), session_flash(p_Item) )
					
				end if
		
			next
		
		end if
	
	next
		
	show_flash_messages = plocal

end function



' ---------------------------------------------------------------
'  Show Static Messages
' ---------------------------------------------------------------
' 
'  @string

function show_static_messages(pType,pText)
	
	show_static_messages = message_box(pType, pText)
	
end function



' ---------------------------------------------------------------
'  Create Messeage Box
' ---------------------------------------------------------------
' 
'  @string

function message_box(pType, pText)
	
	dim plocal
	
	Select Case lcase(pType)
	
		Case "form_error"
			
			plocal = "<div id=""form_error"">" & pText & "</div>"
		
		Case "error"
			
			plocal = "<div id=""errorBox"">" &  pText & "</div>"
		
		Case "warn"
			
			plocal = "<div id=""warnBox"">" &  pText & "</div>"
			
		Case "info"
			
			plocal = "<div id=""infoBox"">" & pText & "</div>"
			
		Case "na"
			
			plocal = pText
		
		Case Else
			
			plocal = "<div id="""&pType&""">" & pText & "</div>"
		
		end select
		
	message_box	= plocal
	
end function 



' ---------------------------------------------------------------
'  Create Form Error Box
' ---------------------------------------------------------------
' 
'  @string

function show_form_errors(pCount, pMessages)
	
	dim local_text, msg
	
	local_text = "<h4>"
	
	if pCount = 1 then
		
		local_text = local_text & "There is "&(pCount)&" error in this form, please check your values</h4>" & Vbcrlf
		
	else
		
		local_text = local_text & "There are "&(pCount)&" errors in this form, please check your values</h4>" & Vbcrlf
		
	end if
	
	
	
	local_text = local_text & "<ul>" & Vbcrlf
	
	for each msg In pMessages
		
		local_text = local_text & error_messages.item(msg)  & Vbcrlf
	
	next
	
	local_text = local_text & "</ul>" & Vbcrlf
	
	show_form_errors = message_box("form_error", local_text)
	
end function
%>