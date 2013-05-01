<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Page:				objPageCache.asp
'Copyright			2002 Joe Audette
'Last Modified:		3/17/2002
'
'You may use this code freely as long as you leave my copyright notice in the code.
'You may modify and extend this code to suit your needs as long as you indicate 
'in your code that it is based on my copyrighted code.
'
'Summary:			This file can be included on any dynamic page and will
'					implement caching of the page for the configured time duration
'					(24 hours by default). 
'					You can also manually force the cache to update by adding 
'					"cachecommand=create" to the query string. The cached 
'					version will be displayed until it expires and is 
'					re-created. 
'					To bypass the cached page and see the dynamic 
'					version, add "cachecommand=bypass" to the query string.
'					To clear the cach for the current page add cachecommand=clear
'					To clear the cach for ALL pages add cachecommand=clearall
'					Note cachecommand=clearall only works when caching to files.
'
'					To use this on a page that posts back to itself you must 
'					include an element within the form like this:
'					<INPUT type="hidden" id="IsPostback" name="IsPostback" value="true">
'					This will bypass the cache and let the dynamic page
'					handle the post back.
'					
'					To set up caching on a page, include this file on the page 
'					and then use code something like this:
'
'					Dim objCache
'					Set objCache = New CPageCache
'
'					**OPTIONAL*************************************
'					if you don't set these properties the cache will 
'					expire in 24 hours
'					objCache.CacheIntervalUnit = CI_MINUTES  
'					'oter choices are CI_HOURS or CI_DAYS
'					objCache.CacheIntervalLength = 10 'how many units
'					***********************************************
'					

'					objCache.AutoCacheToMemory() 'or objCache.AutoCacheToFile()
'					Set objCache = Nothing
'
'					You will also need an Application variable in the global.asa file
'					called Application("CachedContentFolder"), indicating the path to 
'					a folder where the cached pages will be kept when you use AutoCacheToFile. 
'					For example you could put a line in the global.asa file like this:
'					Application("CachedContentFolder") = "c:\CachedContent\"
'					This folder must allow write permission to the IUSER_MACHINENAME user.
'					For best practice this folder should be outside the web site folder
'					structure.
'
'					You will also need an application variable called Application("Root") 
'					indicating the root of your site. For example:
'					Application("Root") = "http://myWEBServer"
'					This helps the object figure out what page its on.
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const	CI_MINUTES = 0
Const	CI_HOURS = 1
Const	CI_DAYS = 1

Class CPageCache

Private mstrCacheFileExt
Private mobjFSO						'FileSystem object
Private mstrCacheFileQualifiers
Private mstrCachedFileName
Private mstrScriptFileName
Private mstrScriptBaseName
Private mstrCachedMemoryName
Private mstrCreateMemoryTime
Private mstrCacheCommand
Private mstrCacheFolder
Private mstrDynURL
Private bIsPostBack
Private mintCacheIntervalUnit
Private mintCacheIntervalLength

dim NEW_FILEPATH

Private Sub Class_Initialize()
	On Error Resume Next
	
	' Create File System Object
	Set mobjFSO = CreateObject("Scripting.FileSystemObject")
	
	'this determines whether to show the cached content or let the page render dynamically
	
	'the default is to pass nothing, in which case the cached page is used
	
	'Set the default cache interval
	mintCacheIntervalUnit = CI_DAYS
	mintCacheIntervalLength = 1
	
	' This is the folder name where the cached content will be stored
	mstrCacheFolder = config_cache_dir
	
	'thes next few lines initialize variables that will determine the cached file
	'name based on the ASP page and any querystring params
	
	NEW_FILEPATH = right(FILEPATH, LEN(FILEPATH) -1)
	
	if NEW_FILEPATH = "" then
		NEW_FILEPATH = "index"
	else 
		NEW_FILEPATH = REPLACE(NEW_FILEPATH, "/","_")
	end if
	
	mstrScriptFileName = server.MapPath(NEW_FILEPATH) 
	mstrScriptBaseName = mobjFSO.GetBaseName(mstrScriptFileName)
	mstrCacheFileQualifiers = GetCacheFileQualifiers()
	mstrCacheFileExt = ".htm" 'this could be any extension
	
	'determine name for output cache file
	mstrCachedFileName = mstrCacheFolder & mstrScriptBaseName & mstrCacheFileExt
	
	mstrCachedFileName = Replace(mstrCachedFileName,"_bypass-cache","")
	mstrCachedFileName = Replace(mstrCachedFileName,"_clear-cache","")
	mstrCachedFileName = Replace(mstrCachedFileName,"_clearall-cache","")
	mstrCachedFileName = Replace(mstrCachedFileName,"_create-cache","")
	
	'If this object is used on a form that posts back to itself
	'include an element within the form like this:
	'<INPUT type="hidden" id="IsPostback" name="IsPostback" value="true">
	'This will bypass the cache and let the dynamic page
	'handle the post back.
	bIsPostBack = false
	If Request("IsPostback") = "true" Then
		bIsPostBack = true
	End If
	
	'get the URL for the current dynamic asp page
	mstrDynURL = config_site_url
	mstrDynURL = mstrDynURL & Cstr(right(FILEPATH, LEN(FILEPATH) -1))
	
	
	if instr(mstrDynURL,"bypass-cache") <> 0 then
	
		mstrCacheCommand = "bypass-cache"
	
	elseif instr(mstrDynURL,"create-cache") <> 0 then
		mstrDynURL = Replace(mstrDynURL,"/create-cache","")
		mstrCacheCommand = "create-cache"
	
	elseif instr(mstrDynURL,"clear-cache") <> 0 then
		mstrDynURL = Replace(mstrDynURL,"/clear-cache","")
		mstrCacheCommand = "clear-cache"
	
	elseif instr(mstrDynURL,"clearall-cache") <> 0 then
		mstrDynURL = Replace(mstrDynURL,"/clearall-cache","")
		mstrCacheCommand = "clearall-cache"
	else
		mstrCacheCommand = ""
	end if
	
	if NOT CACHECONTROL then
		bIsPostBack = true
	end if
	
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Terminate()
	On Error Resume Next
	' Clean up module-level objects
	If Not mobjFSO Is Nothing Then
		Set mobjFSO = Nothing
	End If
	
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PRIVATE MEMBER METHODS
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetCacheFileQualifiers() 'As String
	'this creates a unique file name portion for
	'each unique query string passed to the ASP page
	On Error Resume Next
	Dim strResults
	Dim var
	Dim strSeparator
	strResults = vbNullString
	strSeparator = vbNullString
	For Each var In Request.QueryString
		If var <> "cachecommand" Then
			strResults = strResults & strSeparator & var & "_" & Request.QueryString(var)
			strSeparator = "_"
		End If
	
	Next
	
	GetCacheFileQualifiers = strResults

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CachedFileExists() 'As Boolean
	On Error Resume Next
	CachedFileExists = mobjFSO.FileExists(mstrCachedFileName)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function CacheFileIsExpired() 'As Boolean
	On Error Resume Next
	Dim objFile
	Dim strDate
	Dim strUnit
	
	'Response.write mstrCachedFileName
	'Get the modified date from the file.
	Set objFile = mobjFSO.GetFile(mstrCachedFileName)
	strDate = objFile.DateLastModified
	Set objFile = nothing	
	
	'Determine the cache interval unit that was selected.
	'and set the datediff unit of measurement.
	Select Case mintCacheIntervalUnit
		Case CI_MINUTES
			strUnit = "n"	'minutes
			
		Case CI_HOURS
			strUnit = "h"	'hours
			
		Case CI_DAYS
			strUnit = "d"	'days
			
	End Select
	
	'Return the boolean from the comparison utilizing the datediff function.
	CacheFileIsExpired = (DateDiff(strUnit,CDate(strDate), Now()) > mintCacheIntervalLength)
	
	'Error checking
	If err.Number <> 0 Then
		err.Clear
		CacheFileIsExpired = true
	End If
	
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub RenderCacheFromFile()
	
	On Error Resume Next
	
	Dim objfile, ObjFileProp, FileProp, dead_file
	
	Set ObjFileProp=Server.CreateObject("Scripting.FileSystemObject")
	Set FileProp=ObjFileProp.GetFile(mstrCachedFileName)
	
	Set objfile = mobjFSO.OpenTextFile(mstrCachedFileName,1,false,-1)
	
	'If the cache file is empty handle 
	'the requst.
	if FileProp.Size < 4 then
		Exit Sub
	end  if
	
	'If error happens the dynamic page will handle
	'the request.
	If err.number <> 0 Then
	'	Response.Write err.Description
		Exit Sub
	End If
	
	'Write the cached file to the 
	'Response and end the response.
	Response.write objfile.ReadAll
	Response.end
	
	Set objfile 	= nothing
	Set ObjFileProp	= nothing
	Set FileProp	= nothing
End Sub


' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PUBLIC METHODS
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub AutoCacheToFile()
	On Error Resume Next
	
	If bIsPostBack Then
		'let the dynamic page handle the postback.
		Exit Sub
	End If
	
	Select Case mstrCacheCommand
	
		Case vbNullString 
		
			If Not CachedFileExists() Or CacheFileIsExpired() Then
				CreateCacheFile
			End If
			
			If err.number = 0 Then
				RenderCacheFromFile
			End If

		Case "bypass-cache" 
			'display dynamic page - do nothing
			
		Case "clear-cache"
			ClearCache
		
		Case "clearall-cache"
			ClearAllCachedFiles
			
		Case "create-cache" 
			'create and display cached page
			CreateCacheFile
			If err.number = 0 Then
				RenderCacheFromFile
			End If
				
	End Select
	
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub AutoCacheToMemory()
	On Error Resume Next
	
	If bIsPostBack Then
		'let the dynamic page handle the postback.
		Exit Sub
	End If
	
	Select Case mstrCacheCommand
	
		Case vbNullString 

			If Not CacheExistsInMemory() Or CacheInMemoryIsExpired() Then
				CreateCacheInMemory
			End If
			
			If err.number = 0 Then
				RenderCacheFromMemory
			Else
				'Response.Write err.Description
				'Response.end
				Exit Sub
			End If
			
		Case "bypass-cache" 
			'display dynamic page - do nothing
			
		Case "clear-cache"
			Application(mstrCachedMemoryName) = Null
		
		Case "create-cache" 
			'create and display cached page
			CreateCacheInMemory
			If err.number = 0 Then
				RenderCacheFromMemory
			Else
				'Response.Write err.Description
				'Response.end
				Exit Sub
			End If
			
	End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CreateCacheFile()
	On Error Resume Next
	
	Dim strResponse
	Dim objTextStream
	Dim objHTTP
	Dim strDynURL
	
	'Get the URL of the current ASP request
	strDynURL = mstrDynURL
	
	'Add a querystring param that will bypass the cache
	'forcing a dynamic rendering of the page
	
	strDynURL = strDynURL & "/bypass-cache"
	
	'Instantiate an HTTP object to request the dynamic page
	Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.4.0")
	
	'Depending on your firewall/proxy environment you may need to 
	'configure proxy settings for the HTTP object
	'If your webserver is behind a proxy uncomment the following line.
	'objHTTP.setProxy  SXH_PROXY_SET_PRECONFIG
	'You also need to the use WinHTTP Proxy configuration utility(proxycfg.exe)
	'on the server (see the readme for more details)
	
	'Request the dynamic ASP page
	
	objHTTP.open "GET", strDynURL, false
	objHTTP.send()
	
	'Create a text file using the file system object
	Set objTextStream = mobjFSO.CreateTextFile(mstrCachedFileName, true, true)
	
	'Write the rendered content from the dynamic ASP page to the
	'text file to create the cached content. 
	objTextStream.Write(objHTTP.responseText)
	objTextStream.Close
	
	Set objHTTP = nothing
	Set objTextStream = nothing
	
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ClearCache()
	'Clears the current page from the cache
	On Error Resume Next
	
	If CachedFileExists() Then
		mobjFSO.DeleteFile(mstrCachedFileName)
	End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ClearAllCachedFiles()

	On Error Resume Next
	Dim objFolder
	Dim colFiles
	Dim objFile
	
	'Get a reference to the folder containing cached content files
	Set objFolder = mobjFSO.GetFolder(config_cache_dir)
	Set colFiles = objFolder.Files
	
	'Delete ALL cached pages
	For Each objFile In colFiles
		objFile.Delete()
	Next
	
	Set colFiles = nothing
	Set objFolder = nothing

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CacheIntervalUnit()
	CacheIntervalUnit = mintCacheIntervalUnit
End Property

Public Property Let CacheIntervalUnit(ByVal intNewValue)
	mintCacheIntervalUnit = CInt(intNewValue)
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CacheIntervalLength()
	CacheIntervalLength = mintCacheIntervalLength
End Property

Public Property Let CacheIntervalLength(ByVal intNewValue)
	mintCacheIntervalLength = CInt(intNewValue)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DynURL()
	DynURL = mstrDynURL
End Property

Public Property Let DynURL(ByVal strNewValue)
	mstrDynURL = CStr(strNewValue)
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get StaticFileName()
	CachedFileName = mstrCachedFileName
End Property

Public Property Let CachedFileName(ByVal strNewValue)
	mstrCachedFileName = CStr(strNewValue)
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>