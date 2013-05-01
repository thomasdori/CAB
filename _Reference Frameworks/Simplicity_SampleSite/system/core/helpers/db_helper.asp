<%

' ---------------------------------------------------------------
'  DB Helpers
' ---------------------------------------------------------------
' 
' Connections
' Logging
' Caching


' ---------------------------------------------------------------
'  Database Make Connection
' ---------------------------------------------------------------
' 
' db_make_conn
' db_open
' db_close
' db_system_check
' db_show_name


function db_make_conn(pDsn,pUser,pPwd)
	
	Dim LocalConn
	
	Set LocalConn = Server.CreateObject("ADODB.Connection")
	
	LocalConn.open "dsn=" &pDsn& ";uid=" &pUser& ";pwd=" &pPwd& ";" 
	
	Set db_make_conn = LocalConn
	
end function


' ---------------------------------------------------------------
'  Database Open Connection
' ---------------------------------------------------------------
' 
' 
' 

function db_open(var)
	
	Dim pLocal
	
	pLocal=Trim(var)
	
	Select Case pLocal
	
	Case "dev"
		Set db_open=db_make_conn("local_db","webuser","letmein")
	
	Case "pro"
		Set db_open=db_make_conn("prod_db","webuser","letmein")
		
	End Select
	
	
end function



' ---------------------------------------------------------------
'  Database Close Connection
' ---------------------------------------------------------------
' 
' 
' 

Sub db_close(var)
	var.Close
	Set var = Nothing
End Sub

%>