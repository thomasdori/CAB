<%

	Dim g_ConnectionString
	Dim DBLayer
	Set DBLayer = New DBWrapper
	
	
	Class DBWrapper
	
		Private mConnection
		Private mLastErrorString
		Private mLastErrorNumber    
		
		Private Sub Class_Initialize()
			Set mConnection = Nothing
			g_ConnectionString = ""
		End Sub
			
		Private Sub Class_Terminate()
			Call CloseConnection()
		End Sub
				
		Private Function OpenConnection()
			If mConnection Is Nothing Then
				Set mConnection = CreateObject("ADODB.Connection")
				If g_ConnectionString = "" then
					mConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=isabella;Persist Security Info=True;User ID=sa;Initial Catalog=NorthWind;Data Source=localhost"
				Else
					mConnection.ConnectionString = g_ConnectionString
				End If
				mConnection.Open
			End If
		End Function
		
		Private Function CloseConnection()
			If Not mConnection Is Nothing Then           
			    'If mConnection.State = adStateOpen Then        
			       mConnection.Close                     
			       Set mConnection = Nothing
			    'End If    
			End If		
		End Function

		Public Property Get GetLastErrorString
			GetLastErrorString  = mLastErrorString
		End Property

		Public Property Get  GetLastErrorNumber
			GetLastErrorNumber  = mLastErrorNumber
		End Property
		
		Public Property Get LastStatementFailed
			LastStatementFailed = (GetLastErrorString <> "")
		End Property


		Public Function GetRecordset(sSQL)
			Dim objRs 
			mLastErrorString = ""
			mLastErrorNumber = 0
			Set GetRecordset = Nothing
			
			Call OpenConnection()
			
				Set objRs = CreateObject("ADODB.Recordset")			
			
				objRs.CursorLocation = 3 'adUseClient
				objRs.CursorType     = 0 'adOpenForwardOnly
				objRs.LockType       = 4 'adLockBatchOptimistic
			
				On Error Resume Next	
				objRs.Open  sSQL, mConnection                    		
	
				If Err.number <> 0 Then					
					mLastErrorString = Err.Description
					mLastErrorNumber = Err.number					
					Page.DebugEnabled = True
					'Page.DebugLevel = 3
					Page.TraceError Page.Control,"DBLayer Error SQL:" & sSQL,10										
					Err.Clear					
					CloseConnection()
					Exit Function
				End If
			
				On Error Goto 0
				
				If objRs.RecordCount > 0 Then
					objRs.MoveFirst                                  
				End If
			
				Set objRs.ActiveConnection = Nothing					
				Set GetRecordset = objRs	

			Call CloseConnection()
		End Function
	
		Public Function UpdateRecordSet(rs)
			Call OpenConnection()
				Set rs.ActiveConnection = mConnection
				rs.UpdateBatch
			Call CloseConnection()
		End Function

		Public Function ExecuteSQL(Sql)
			'Dim rows
			Call OpenConnection()
				mConnection.Execute Sql ', rows				
			Call CloseConnection()
		End Function
	
		Public Function GetPagedRecordSet(sSelect,sFrom, sWhere,sSort,iPageSize,iPageIndex,iNumRows)
			Dim sSQL
			Dim sSQLTemplate
			Dim S2
			Dim rs
			
			S2 = Replace(sSort," DESC"," _S_")
			S2 = Replace(S2," ASC"," DESC")
			S2 = Replace(S2," _S_"," ASC")
			
			sSQLTemplate = "SELECT * FROM (SELECT TOP {PS} * FROM  (SELECT TOP {PS2} {C} FROM {F} {W} ORDER BY {S1}) as T1 ORDER BY {S2}) AS T2 ORDER BY {S1}  "
			
			sSQLTemplate = Replace(sSQLTemplate,"{S2}",S2,1,1)
			sSQLTemplate = Replace(sSQLTemplate,"{S1}",sSort,1,2)
			sSQLTemplate = Replace(sSQLTemplate,"{C}",sSelect,1)
			sSQLTemplate = Replace(sSQLTemplate,"{F}",sFrom,1)
			
			If sWhere<>"" Then
				sSQLTemplate = Replace(sSQLTemplate,"{W}","WHERE " & sWhere,1)
			Else
				sSQLTemplate = Replace(sSQLTemplate,"{W}","",1)
			End If
			
			sSQLTemplate = Replace(sSQLTemplate,"{PS}",iPageSize,1)
			sSQLTemplate = Replace(sSQLTemplate,"{PS2}",iPageSize*(1 + iPageIndex),1)
			
			If iPageIndex = 0 And iNumRows=0 Then
				sSQL = "SELECT Count(*) FROM " & sFROM 
				If sWhere<>"" Then
					sSQL = sSQL & " WHERE " & sWHERE
				End If
				Set rs = GetRecordSet(sSQL)
				iNumRows = rs(0).Value
				Set rs = Nothing
			End If			
			
			Set GetPagedRecordSet = GetRecordSet(sSQLTemplate)			
		End Function

	End Class


	Class DBStoredProcedure
		Dim Name
		Dim Parameters

		Private Sub Class_Initialize()
			Set Parameters = New SPParameters
		End Sub

		Private Sub Class_Terminate()
			Set Parameters = Nothing
		End Sub
		
		Public Function Execute()
			Dim sSQL
			
			sSQL = "exec " & Name  & " " & Parameters
			
			Set Execute =  DBLayer.GetRecordSet(sSQL)				
			
		End Function
		
		Public Function ExecuteScalar() 
			Dim sSQL
			Dim rs

			sSQL = "exec " & Name  & " " & Parameters
			Set rs =  DBLayer.GetRecordset(sSQL)
			
			If Not rs Is Nothing Then
				If rs.RecordCount > 0 Then
					ExecuteScalar = rs(0).Value
				End If
			
				rs.Close
				Set rs = Nothing		
			End If
						
		End Function

		Public Function ExecuteNoReturn()
			Dim sSQL
			
			sSQL = "exec " & Name  & " " & Parameters
			DBLayer.ExecuteSQL(sSQL)				
			
		End Function

		
	End Class

	Class SPParameters
		Private mParameters
		
		Private Sub Class_Initialize()
			Set mParameters = New StringBuilder
		End Sub
			
		Private Sub Class_Terminate()
			Set mParameters = Nothing
		End Sub
		
		
		Public Function AddString(Name,Value)
			If IsNull(Value) Then
				mParameters.Append "@" & Name & "=NULL"
			Else
				mParameters.Append "@" & Name & "='" & Replace(Value,"'","''") & "'"
			End If
		End Function
	
		Public Function AddNumeric(Name,Value)
			mParameters.Append "@" & Name & "=" & IIF(IsNull(Value),"NULL",Value)
		End Function
	
		Public Function AddDate(Name,Value)
			If IsNull(Value) Or "" & Value = "" Then
				mParameters.Append "@" & Name & "=NULL"
			Else
				mParameters.Append "@" & Name & "='" & Replace(Value,"'","''") & "'"
			End If
		End Function
		
		Public Default Function GetParametersText()
				GetParametersText = mParameters.GetJoined(",")
		End Function
		
	End Class

	Public Function GetRecordset(sSQL)
		Set GetRecordset = DBLayer.GetRecordset(sSQL)
	End Function
	
%>