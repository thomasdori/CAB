<%

' ---------------------------------------------------------------
' 	Server Variables Helper
' ---------------------------------------------------------------
' 
' sv_hostip
' sv_useragent
' sv_get_value

' ---------------------------------------------------------------
'  Server Variables Host IP
' ---------------------------------------------------------------
' 
'  @string

function sv_hostip
	
	dim local_str
	
	local_str = sv_get_value("REMOTE_ADDR")
	
	sv_hostip =  local_str
	
end function



' ---------------------------------------------------------------
'  Server Variables User Agent
' ---------------------------------------------------------------
' 
'  @string

function sv_useragent
	
	dim local_str
	
	local_str = sv_get_value("HTTP_USER_AGENT")
	
	local_str = Left(local_str,200)
	
	sv_useragent =  local_str
	
end function



' ---------------------------------------------------------------
'  Server Variables Get Value
' ---------------------------------------------------------------
' 
'  @string

function sv_get_value(var)
	
	sv_get_value = Request.ServerVariables(var)
	
end function
%>