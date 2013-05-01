<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server File Uploader
	
	'Helper function.
	Public Function New_ServerFileUploader(name,path) 
		Set New_ServerFileUploader = New ServerFileUploader
		New_ServerFileUploader.Control.Name = name
		New_ServerFileUploader.TempUploadPath = path
	End Function

	'JavaScript Includes
	Dim ServerFileUploader_jsRegistered

	Class ServerFileUploader
		Public Control
		Public TempUploadPath
		Public FileName
		Public FileContentType
		Public Text		
		Public FileFilter
		Public MaxUploadSize

		Private mFileId
		Private mFileSize
		Private mTmpFileName		

		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Control.ImplementsProcessPostBack = True

			TempUploadPath	= "./" ' defaults to current root
			FileName	= ""
			FileFilter	= ""
			MaxUploadSize = 0
			mFileSize	= 0
			mFileId = ""			
		End Sub

		Private Sub Class_Terminate()
			'save special variables to session for file upload
			Session(Control.Name+"_FileID")			= mFileId
			Session(Control.Name+"_TmpUploadPath")	= TempUploadPath
			Session(Control.Name+"_FileFilter") 	= FileFilter
			Session(Control.Name+"_MaxUploadSize")	= MaxUploadSize
		End Sub
	     
		Public Function ReadProperties(bag)			
	   		Text			= bag.Read("T")
	   		FileFilter		= bag.Read("F")
	   		MaxUploadSize	= bag.Read("M")
	   		TempUploadPath	= bag.Read("P")
	   		mFileId			= bag.Read("I")	   		
	   		If Len(FileName)=0 Then 
	   			Call ProcessPostBack()
	   		End If
		End Function
		
		Public Function WriteProperties(bag)
	   		If Len(mFileId) = 0 Then 
	   			' create temp file id
	   			mFileId = Control.Name + "." + Session.SessionID
	   		End If
	   		bag.Write "T",Text
	   		bag.Write "I",mFileId
	   		bag.Write "F", FileFilter
	   		bag.Write "M",MaxUploadSize
	   		bag.Write "P", TempUploadPath
	End Function

		Public Function ProcessPostBack()
			FileName		= Session(Control.Name+"_FileName")
			FileContentType	= Session(Control.Name+"_FileCType")
			mFileSize		= Session(mControlName+"_FileSize")			
		End Function
	   
	   Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunction(e.EventFnc)
	   End Function			

		Public Property Get FileSize()
			FileSize = mFileSize
		End Property
		
		Public Property Get TempFileName()
			Dim tPath
			TempFileName = mFileID
		End Property		
	   
		Public Function SaveFile(sPath)
			Dim f
			Dim fso
			Dim tPath

			tPath = TempUploadPath
			If tPath = "" Then Exit Function
			If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
			If Mid(tPath, Len(tPath)) <> "\" Then tPath = tPath & "\"
					
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If FileName<>"" And fso.FolderExists(sPath) Then
				If fso.FileExists(tPath + mFileId) Then 
					Set f = fso.GetFile(tPath + mFileId)
					If fso.FileExists(spath+FileName) Then  fso.DeleteFile(spath+FileName)
					f.Move(sPath + FileName)
				End If
			End If			
		End Function

		Public Sub DiscardFile()
			Dim fso
			Dim tPath

			tPath = TempUploadPath
			If tPath = "" Then Exit Sub
			If Mid(tPath, Len(tPath)) <> "\" Then tPath = tPath & "\"

			' remove temporary file if it exists
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If fso.FolderExists(tPath) Then
				If fso.FileExists(tPath + mFileId) Then 
					fso.DeleteFile(tPath + mFileId)
				End If
			End If
		End Sub
	   
		Public Default Function Render()

			Dim varStart	 
			
			If Control.Visible = False Then
				Exit Function
			End If			 

			varStart = Now

			If Not ServerFileUploader_jsRegistered Then 				
				ServerFileUploader_jsRegistered = True ' file upload script
				Page.RegisterClientScript "Calendar","<script language='JavaScript'>" + vbcrlf + _
					"function showFileWindow(id){"+ vbcrlf + _
					"	var l = (screen.width-250)/2;"+ vbcrlf + _
					"	var t = (screen.height-60)/2;"+ vbcrlf + _
					"	var url = SCRIPT_LIBRARY_PATH + 'fileuploader/upload.asp?id='+id;"+ vbcrlf + _
					"	var win = window.open(url,'file_uploader'," + vbcrlf + _
					"	'top='+ t +',left='+l+',width=250,height=80,status=yes,resizable=no,dependent=yes,alwaysRaised=yes'"+ vbcrlf + _
					"	);"+ vbcrlf + _
					"	if(!win) alert('Unable to open window. Please configure your popup blocker to allow popups for this site.');"+ vbcrlf + _
					"	else {"+ vbcrlf + _
					"		win.focus();"+ vbcrlf + _
					"	}"+ vbcrlf + _
					"};"+ vbcrlf + _
					"</script>"
			End If

			Response.Write  "<input type='button' id='" & Control.Name & "' name='" &_
					Control.Name & "' value='" & Server.HTMLEncode(Text) & "' "  &_
					" tabindex = " & Control.TabIndex
					If Control.CssClass	 <> "" Then Response.Write " class='" & Control.CssClass  & "' "
					If Control.Style		 <> "" Then Response.Write " style='" & Control.Style	  & "' "
					If Control.Attributes <> "" Then Response.Write Control.Attributes   & " "
					If Not Control.Enabled   Then Response.Write " Disabled  "
					Response.Write " onclick='showFileWindow(""" & Control.Name &""")'>"
			Page.TraceRender varStart,Now,Control.Name

		End Function

	End Class


'::File Upload Functions::::::::::::::::::::::::::::::::::::::::::::::
'::
'::This section is only executed during file upload
'::to handle the multipart document encoded format


If Request.QueryString("UpLoadSave") = "on" Then
	Dim mFileId
	Dim mUploadPath
	Dim mControlName
	Dim mFile, mFileUploader 
	Dim mFileFilter,mMaxUploadSize
	Dim mFileUploadMsg
	Dim mFilArray
	Dim mFileOk
	Dim mFileUploadLink

	' Create the FileUploader
	Set mFileUploader = New FileUploader	
	
	' get file settings
	mFileUploader.Upload()
	mControlName	= mFileUploader.Form("ctrl")	'get control name
	mFileId			= Session(mControlName+"_FileId")
	mUploadPath 	= Session(mControlName+"_TmpUploadPath")
	mFileFilter		= Session(mControlName+"_FileFilter")
	mMaxUploadSize	= Session(mControlName+"_MaxUploadSize")
	
	'loop through file objects	
	For Each mFile  in  mFileUploader.Files.Items

		mFileOk = True	


		'check filters
		If mFileOk And Len(mFileFilter)>0 Then
			mFilArray = Split(mFile.FileName,".") 'get file extension
			If UBound(mFilArray)>0 Then
				If Instr(1,LCase(","& mFileFilter &","),LCase(","& mFilArray(1) &","))>0 Then 
					mFileOk = True
				Else
					mFileOk = False
					mFileUploadMsg = "Unable to upload '."+mFilArray(1)+"' files."
				End If
			End If
		End If

		'check max size
		If mFileOk And mMaxUploadSize >0 Then
			If mFile.FileSize <= mMaxUploadSize Then
				mFileOk = True
			Else
				mFileOk = False
				mFileUploadMsg = "Unable to upload files larger than "& mMaxUploadSize &" bytes"
			End If

		End If
		Exit For
	Next 	
	If mFileOk Then
		' save file name, content type and size 
		mFileUploadMsg = "File successfully uploaded."
		mFileUploadLink = "<a href='javascript:RaiseUploadEvent()' style='text-decoration:none;color:blue;'>Close This Window</a>" + _ 
		"<script>window.setTimeout('RaiseUploadEvent()',700)</script>" 
		Session(mControlName+"_FileName") = mFile.FileName
		Session(mControlName+"_FileSize") = mFile.FileSize
		Session(mControlName+"_FileCType") = mFile.ContentType
		mFile.FileName = mFileId
		mFile.SaveToDisk mUploadPath
	Else
		mFileUploadLink = "<a href='javascript:history.go(-1)' style='text-decoration:none;color:blue;'>&lt; Back</a>"
	End If
	
	%>
	<html>
	<head>
	<title>File Uploader</title>
	<script>
		function RaiseUploadEvent(){
			window.opener.doPostBack('OnUpload','<%=mControlName%>');
			window.close();
		}
	</script>
	<style>
	TD {
		font-family:verdana,arial;
		font-size:11px;
	}
	</style>	
	</head>
	<body bgcolor="#ECE9D8" topmargin="0">
	<table width="100%" border="0">
	<tr><td><font size="2" face="arial"><%=mFileUploadLink%></font></td></tr>
	<tr><td><hr size="1"></td></tr>
	<tr><td align="center">
	 	<b><font size="4" face="arial"><%=mFileUploadMsg%></font></b><br>
	</td></tr></table>
	</body>
	</html>
	<%
End If	


'***************************************
' File:	  Upload.asp
' Author: Jacob "Beezle" Gilley
' Email:  avis7@airmail.net
' Date:   12/07/2000
' Comments: The code for the Upload, CByteString, 
'			CWideString	subroutines was originally 
'			written by Philippe Collignon...or so 
'			he claims. Also, I am not responsible
'			for any ill effects this script may
'			cause and provide this script "AS IS".
'			Enjoy!
'****************************************

Class FileUploader
	Public  Files
	Private mcolFormElem

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		End If
		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		End If
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default Sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
		Dim nPosFile, nPosBound

		biData = Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
		
		If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
		 
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)
		
		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
			
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile
				
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				
				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				
				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
				
				If oUploadFile.FileSize > 0 Then Files.Add LCase(sInputName), oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			End If

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public Sub SaveToDisk(sPath)
		Dim oFS, oFile
		Dim nIndex
	
		If sPath = "" Or FileName = "" Then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
	
		Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		If Not oFS.FolderExists(sPath) Then Exit Sub
		
		Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		
		For nIndex = 1 to LenB(FileData)
		    oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
		Next

		oFile.Close
	End Sub

End Class


%>