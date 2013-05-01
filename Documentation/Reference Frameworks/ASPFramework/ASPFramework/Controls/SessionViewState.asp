<%
	'THIS VIEWSTATE MODE IS USEFUL IF YOU WANT TO COMPRESS THE VIEWSTATE IN A SESSION VARIABLE.
	'REQUIRES ASPFramework.dll
	Public Function Page_LoadPageStateFromPersistenceMedium()
		Dim obj
		Page.TraceMsg "Page_LoadPageStateFromPersistenceMedium","SessionViewState"
		Set obj = CreateObject("ASPFramework.Crypto")		
			Page.ViewState.LoadViewState( obj.UncompressString(  Session("VS2")  ) )
			Session("CLASP_CVWS") = ""
		Set obj = Nothing
	End Function
	
	Public Function Page_SavePageStateToPersistenceMedium()
		Dim obj
		Page.TraceMsg "Page_SavePageStateToPersistenceMedium","SessionViewState"
		Set obj = CreateObject("ASPFramework.Crypto")		
			Session("CLASP_CVWS") = obj.CompressString( Page.ViewState.GetViewState() )
		Set obj = Nothing
	End Function

%>