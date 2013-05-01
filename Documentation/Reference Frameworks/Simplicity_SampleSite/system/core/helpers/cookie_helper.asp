<%

' ---------------------------------------------------------------
'  Cookie Helper
' ---------------------------------------------------------------
' 
' cookie_read
' cookie_write
' cookie_remove


' ---------------------------------------------------------------
'  Cookie Read
' ---------------------------------------------------------------
' 
'  @string

function cookie_read(var)
	
	Dim pLocal
	
	pLocal		=  Request.Cookies(config_cookie_domain)(var)
	
	cookie_read = pLocal
	
end function



' ---------------------------------------------------------------
'  Cookie Write
' ---------------------------------------------------------------
' 
' 

function cookie_write(var,var2)
	
	Response.Cookies(config_cookie_domain).Domain=config_cookie_url
	
	Response.Cookies(config_cookie_domain)(var)=var2
	
	Response.Cookies(config_cookie_domain).Expires=config_cookie_expire
	
end function


' ---------------------------------------------------------------
'  Cookie Write
' ---------------------------------------------------------------
' 
' 

function cookie_remove(var)
	
	Response.Cookies(config_cookie_domain).Domain=config_cookie_url
	
	Response.Cookies(config_cookie_domain)(var)= ""
	
	Response.Cookies(config_cookie_domain).Expires= date - 100
	
end function



%>