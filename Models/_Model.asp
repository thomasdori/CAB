<%
	Dim model : Set model = New ModelClass

	Class ModelClass
		Public Function CallProcedure(procedureName, parameters)




		Dim recordSet : Set recordSet = Server.CreateObject("ADODB.Recordset")
		Dim objConn   : Set objConn	  = Server.CreateObject("ADODB.Connection")
		Dim command   : Set command   = Server.CreateObject("ADODB.Command")

		objConn.ConnectionString = GetConnectionString()
		objConn.Open

		command.ActiveConnection = con
		command.CommandType = adobjConnStoredProc
		command.CommandTimeout = 0
		command.CommandText = procedureName
		command.Parameters.Append objConn.CreateParameter("@usNo", adInteger, adParamInput, -1, session("usrNr"))

		Set recordSet = command.execute()

		With recordSet
			.CursorLocation = 3 'adUseClient
			.CursorType     = 0 'adOpenForwardOnly
			.LockType       = 4 'adLockBatchOptimistic
		End With

		Do While Not recordSet.Eof
			' do something with recordSet("column")
			recordSet.movenext
		Loop

	    recordSet.Close
	    Set recordSet = Nothing






			Set objConn = Nothing

			If recordSet.RecordCount > 0 Then

			End If



		End Function

		Private Function GetConnectionString()
			GetConnectionString = Session("connectionString")
		End Function
	End Class
%>