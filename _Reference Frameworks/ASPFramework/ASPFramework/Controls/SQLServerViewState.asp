<%
'drop PROCEDURE CLASP_UpdateViewState 
'go
'CREATE PROCEDURE CLASP_UpdateViewState (@SessionID char(32) , @VsData  IMAGE) AS
'
'IF NOT (EXISTS(SELECT @SessionID FROM CLASPViewState WHERE SessionID = @SessionID) )
'		
'	INSERT INTO CLASPViewState (SessionID,SessionItemLong) VALUES(@SessionID,@VsData)
'
'ELSE
'	UPDATE CLASPViewState SET SessionItemLong = @VsData, Updated = getutcdate() WHERE SessionID = @SessionID
'
'go
'drop PROCEDURE CLASP_GetViewState
'go
'CREATE PROCEDURE CLASP_GetViewState (@SessionID char(32) )AS
'SELECT SessionID,SessionItemLong,Updated FROM CLASPViewState WHERE SessionID = @SessionID

	Dim SQLViewStateRs
	Set SQLViewStateRs = Nothing
	Page.ViewStateMode = VIEW_STATE_MODE_SQL_SERVER
	Public Function Page_LoadPageStateFromPersistenceMedium()
			Dim objConn			

			Page.TraceMsg "Page_LoadPageStateFromPersistenceMedium","SQLServerViewState"

			Set SQLViewStateRs	= CreateObject("ADODB.Recordset")				
			Set objConn			= CreateObject("ADODB.Connection")
			With SQLViewStateRs
				.CursorLocation = 3 'adUseClient
				.CursorType     = 0 'adOpenForwardOnly
				.LockType       = 4 'adLockBatchOptimistic			
			End With
			
			With objConn
				.ConnectionString = "Provider=SQLOLEDB.1;Password=isabella;Persist Security Info=True;User ID=sa;Initial Catalog=EKIDb;Data Source=localhost"
				.Open
				SQLViewStateRs.Open "CLASP_GetViewState '" & Session.SessionID & "'",objConn
				Set SQLViewStateRs.ActiveConnection = Nothing
				.Close				
			End With
			
			Set objConn = Nothing
			
			If SQLViewStateRs.RecordCount > 0 Then							
				Page.ViewState.LoadViewState SQLViewStateRs.Fields("SessionItemLong").Value
			Else
				Page.ViewState.Clear
			End If
			
			
	End Function
	
	Public Function Page_SavePageStateToPersistenceMedium()

			Dim objConn								
			Set objConn	= CreateObject("ADODB.Connection")

			Page.TraceMsg "Page_SavePageStateToPersistenceMedium","SQLServerViewState"
						
			objConn.ConnectionString = "Provider=SQLOLEDB.1;Password=isabella;Persist Security Info=True;User ID=sa;Initial Catalog=EKIDb;Data Source=localhost"
			objConn.Open

			If SQLViewStateRs Is Nothing Then				
				Set SQLViewStateRs	= CreateObject("ADODB.Recordset")					
				With SQLViewStateRs
					.CursorLocation = 3 'adUseClient
					.CursorType     = 0 'adOpenForwardOnly
					.LockType       = 4 'adLockBatchOptimistic			
				End With
				SQLViewStateRs.Open "CLASP_GetViewState '" & Session.SessionID & "'",objConn
			Else
				Set SQLViewStateRs.ActiveConnection = objConn
			End If
				
			If SQLViewStateRs.RecordCount = 0 Then
				SQLViewStateRs.AddNew
			End If	
			
			SQLViewStateRs.Fields("SessionID").Value = Session.SessionID
			SQLViewStateRs.Fields("SessionItemLong").AppendChunk  Page.ViewState.GetViewState
			SQLViewStateRs.Fields("Updated").Value = Now
			SQLViewStateRs.UpdateBatch									
			
			SQLViewStateRs.Close				
			objConn.Close
			Set SQLViewStateRs = Nothing
			Set objConn = Nothing			
		
	End Function

%>