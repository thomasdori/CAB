<!-- #include File="../Models/ErrorModel.asp" -->

<%
	Dim logging : Set logging = New LoggingHelperClass

	Class LoggingHelperClass
		Private Sub Log(message)
			'todo: decide if you want to log to the response or to the database
		End Sub

		Public Sub Log(message, errorLevel)
			If errorLevel <= LOG_LEVEL
				Log(message)
			End If
		End Sub
	End Class
%>