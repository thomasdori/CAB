<!-- #include File="../Libraries/adovbs.inc" -->
<!-- #include File="../Extensions/List.asp" -->

<%
	Dim model : Set model = New ModelClass

	Class ModelClass
		Public Function CallProcedure(procedureName, parameterList, hasContentHandler, rowHandler)
			Dim recordSet 	  : Set recordSet 	  = Server.CreateObject("ADODB.Recordset")
			Dim objConn   	  : Set objConn	  	  = Server.CreateObject("ADODB.Connection")
			Dim command   	  : Set command   	  = Server.CreateObject("ADODB.Command")

			objConn.ConnectionString = GetConnectionString()
			objConn.Open

			command.ActiveConnection = con
			command.CommandType = adobjConnStoredProc
			command.CommandTimeout = 0
			command.CommandText = procedureName


			'todo: iterate over parameter
			For Each parameter In parameterList
				command.Parameters.Append objConn.CreateParameter(parameter.Name, parameter.Type, adParamInput, -1, parameter.Value)
			Next

			Set recordSet = command.Execute()

			With recordSet
				.CursorLocation = 3 'adUseClient
				.CursorType     = 0 'adOpenForwardOnly
				.LockType       = 4 'adLockBatchOptimistic
			End With

			If Not recordSet.Eof
				Call HasContentHandler(recordSet)
				Do While Not recordSet.Eof
					Call RowHandler(recordSet)
					recordSet.movenext
				Loop
			End If

		    recordSet.Close
		    Set recordSet = Nothing

		    objConn.Close
			Set objConn = Nothing
		End Function

		Private Function AddParameter(paremeter)

		End Function

		Private Function GetConnectionString()
			GetConnectionString = Session("connectionString")
		End Function
	End Class
%>