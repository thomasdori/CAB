<%	'You may need to create the viewstate folder (see VS_FILE_PATH)
	
	Dim VS_FILE_PATH
	VS_FILE_PATH = "\ASPFrameWork\ViewState\"
	Public Function Page_LoadPageStateFromPersistenceMedium()
			Dim oFSO
			Dim oTS
			Dim sFile
			
			Page.TraceMsg "Page_LoadPageStateFromPersistenceMedium","FileViewState"
			
			If Session("CLASP_FVS") = "" Then
				sFile = Server.MapPath(VS_FILE_PATH & Session.SessionID)
				Session("CLASP_FVS") = sFile
			Else
				sFile = Session("CLASP_FVS")
			End If
			
			Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
			Set oTS = oFSO.OpenTextFile(sFile,1,False) '1 = ForReading
			
			Page.ViewState.LoadViewState oTS.ReadAll
			oTs.Close
			
			Set oTS = Nothing
			Set oFSO = Nothing
	End Function
	
	Public Function Page_SavePageStateToPersistenceMedium()
			Dim oFSO
			Dim oTS
			Dim sFile
			Page.TraceMsg "Page_SavePageStateToPersistenceMedium","FileViewState"
			
			If Session("CLASP_FVS") = "" Then
				sFile = Server.MapPath(VS_FILE_PATH & Session.SessionID)
				Session("CLASP_FVS") = sFile
			Else
				sFile = Session("CLASP_FVS")
			End If

			Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
			
			'AUTO CLEAN UP SECTION
				If Application("CLASP_FVS_CU") <> Day(Now) Then
					Application.Lock
						Application("CLASP_FVS_CU") = Day(Now)
						Dim oFile
						Dim oFolder
						Set oFolder = OFSO.GetFolder(Server.MapPath(VS_FILE_PATH))
						For Each oFile In oFolder.Files
							If DateDiff("n",oFile.DateLastModified,Now) > 60 Then
								oFile.Delete True
							End If
						Next
						Set oFolder = Nothing
						Set oFile   = Nothing
					Application.UnLock
				End If
			'END AUTO
			
			Set oTS = oFSO.OpenTextFile(sFile,2,True) '2 = ForWriting			

			oTS.Write Page.ViewState.GetViewState
			oTs.Close
			Set oTS = Nothing
			Set oFSO = Nothing
			
	End Function


%>
