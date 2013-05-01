<%

' ---------------------------------------------------------------
'  Session Helpers
' ---------------------------------------------------------------
' 
' session_check
' session_read
' session_write
' session_set_flash
' session_flash
' session_remove



' ---------------------------------------------------------------
'  Session Check for Active Session
' ---------------------------------------------------------------
' 
' 

function session_check(var)

end function


' ---------------------------------------------------------------
'  Read a Value from a session
' ---------------------------------------------------------------
' 
' 

function session_read(var)
	
	session_read = Session.Contents(var)
	
end function


' ---------------------------------------------------------------
'  Write a Value to a session
' ---------------------------------------------------------------
' 
' 

sub session_write(var,data)
	
	Session.Contents(var) = data
	
end sub


' ---------------------------------------------------------------
'  Set A Flash Message
' ---------------------------------------------------------------
' 
' 

sub session_set_flash(var, data)
	
	Call cookie_write(var & "_flash",data)
	
end sub


' ---------------------------------------------------------------
'  Read A Flash Message, then remove the session paramater
' ---------------------------------------------------------------
' 
' 

function session_flash(var)
	
	dim session_name
	
	session_name = cookie_read( var )
	
	if does_exists(session_name) then
		
		Call cookie_remove( var )
		
	end if
	
	session_flash = session_name
	
end function


' ---------------------------------------------------------------
'  Remove a Session parameter 
' ---------------------------------------------------------------
' 
' 

sub session_remove(var)
	
	Session.Contents.Remove(var)
	
end sub


' ---------------------------------------------------------------
'  Remove All Session parameters
' ---------------------------------------------------------------
' 
' 

sub session_remove_all
	
	Session.Contents.RemoveAll()
	
end sub


%>
