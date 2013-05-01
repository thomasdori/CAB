<%

Class CallStackTrace
	
	Private mCurrentLevel	
	Private mGlobalTrace
	
	Private Sub Class_Initialize()
		
		'Defaults to 15 nested calls		 
		 mCurrentLevel = 0
		 Set mGlobalTrace = New StringBuilder
	End Sub
	
	
	Public Sub BeginTrace(name)		
		mGlobalTrace.Append Now & String(mCurrentLevel*5,"-") & " START: " & name & "<BR>"
		mCurrentLevel = mCurrentLevel + 1
	End Sub
	
	Public Sub Trace(msg)		
			mGlobalTrace.Append Now & String(mCurrentLevel*5,"-") & msg & "<BR>"
	End Sub
	
	Public Sub EndTrace()
			mCurrentLevel = mCurrentLevel - 1
			mGlobalTrace.Append Now & String(mCurrentLevel*5,"-") & " END <BR>"			
	End Sub
	
	Public Function ToString()
		ToString  = mGlobalTrace.ToString()
	End Function
	


End Class


%>