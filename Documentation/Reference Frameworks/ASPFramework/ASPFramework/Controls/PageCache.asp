<%
Dim PageCache 
Set PageCache = New ASPCache
Call PageCache.ResolveCache()

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::CACHE OBJECT
		
Class ASPCache
	Dim CacheDuration
	Dim VaryByParams
	Dim ByPassPostBack
	Dim CacheToMemory
	Private mScriptName
	
	Private Sub Class_Initialize()
		CacheDuration = 60 '60 minutes
		VaryByParams  = False
		ByPassPostBack = True 'Don't use cache on postbacks
		CacheToMemory  = True			
	End Sub
	
	Public Property Get ScriptName
		
		If mScriptName = "" Then
			mScriptName = "_CACHE" & Request.ServerVariables("PATH_INFO")
			mScriptName = Replace(mScriptName, "/","_")
			mScriptName = Replace(mScriptName, "\","_")
			'If VaryByParams Then
			'	If Request.QueryString <> "" Then
			'		mScriptName = Replace(mScriptName, "\","_")
			'	End If
			'End If
		End If			
		ScriptName = mScriptName
	End Property
	
	Private Sub Class_Terminate()			
	End Sub
	
	Public Function RequestCacheAction()
		Dim fnc, bCacheYN, iCacheDuration
		RequestCacheAction = False
		Set fnc = GetFunctionReference("Page_RequestCacheAction")		
		If Not fnc Is Nothing Then
			RequestCacheAction = True			
			Call fnc(bCacheYN)				
			RequestCacheAction = bCacheYN
		End If			
	End Function

	
	Public Function FlushCache()
		Dim x,varKey
		Dim obj
			
		Set obj = Application.Contents
							
		If obj.Count = 0  Then
			Exit Function
		End If
			
		For x = 1 To obj.Count
			varKey = obj.Key(x)
			If Left(varKey,6) = "_CACHE"  Then
				obj.Item(varKey) = ""
			End If
		Next						

	End Function
	
	Public Function ResolveCache()
		
		Dim CacheCmd
		Dim oXmlHttp
		
		ResolveCache = False
		
		'Checks if the page should be cached Y/N
		If Not RequestCacheAction Then
			Exit Function
		Else
			'Handle at least security!
			Page.HandleServerEvent "Page_Authenticate_Request"
			Page.HandleServerEvent "Page_Authorize_Request"
		End If
		
		If Page.IsPostBack And ByPassPostBack Then
			Exit Function
		End If
		
		CacheCmd = Request.QueryString("cachecommand")
								
		If CacheCmd = "refresh" Or cacheCmd = "create" Or CacheCmd = "delete" Then
			Page.TraceMsg "ResolveCache", CacheCmd & " cache"
			Application(ScriptName) = ""
			Application(ScriptName & "_Date") = ""
		End If
		
		If CacheCmd = "flush" Then
			Call FlushCache() 'oh oh
		End If
		
		'If Creating or Refreshing then just let it go...
		If Not (CacheCmd = "refresh" Or cacheCmd = "create" Or CacheCmd = "delete") Then						
			If Application(ScriptName) = "" Then
				'Pass the querystring
				Set oXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
													
				oXmlHttp.setTimeouts 1500, 1500, 1500,5000
				oXmlHttp.open "GET", "http://127.0.0.1" & Request.ServerVariables("PATH_INFO") & "?cachecommand=create"
				oXmlHttp.send
				Application(ScriptName)=oXmlHttp.responseText
				Application(ScriptName & "_Date" )= DateAdd("n",CacheDuration,Now)
				Set oXmlHttp = Nothing
				ResolveCache = True
			Else					
				Response.Clear					
				'Response.Write Application(ScriptName & "_Date") & "<BR>"
				'Response.Write PageCache.CacheDuration & "<BR>"					
				'Response.Write  DateDiff("n",Now,CDate(Application(ScriptName & "_Date"))) & "<BR>"
				
				Response.Write Application(ScriptName)
				
				'Expire the cache if needed!
				If DateDiff("n",Now,CDate(Application(ScriptName & "_Date"))) <= 0 Then
					Application(ScriptName)	 = ""
				End If					
				Response.End					
			End If
		End If
		'You have to!, otherwise the page will not have a viewstate to restore to when
		'loading from the cache!...
		Page.ViewStateMode = VIEW_STATE_MODE_CLIENT
	End Function

End Class
%>