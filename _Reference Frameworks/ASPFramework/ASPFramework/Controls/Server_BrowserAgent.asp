<%

	Class BrowserAgent
		Public IsDOM
		Public IsGecko
		Public IsOpera
		Public IsMSIE
		Public IsSafari
		Public IsNS
		Public IsMAC
		Public IsWin32
		Public IsUnix
		Public HasNoScript
		Public HasNoIFRAME
		Public ScreenWidth
		Public ScreenHeight
		Public ScreenColor
		
		Public IsNoScript 'deprecated
		
		Private mVer
		Private mApp
		Private mUa 
		Private mScr

		Private Sub Class_Initialize()

			'# This check will only be executed once when the first page is loaded
			If Session("BrowserAgent") = "" Then
				On Error Resume Next	
				Server.Transfer WEB_CONTROLS_PATH+"Server_BrowserCapture.asp"
				'if we get here then there was an error
				Response.Write "<font face='arial' size='3'><b>PATH OR ACCESS ERROR:</b><hr> Unable to load Server_BrowserCapture.asp. Please ensure that you've correctly set the CLASP Path inside the CLASP_Setup.asp file.</font>"
				Response.Write "<blockquote><font face='arial' size='3'> Current Paths:<br><blockquote>"+WEB_CONTROLS_PATH+"<br>"+SCRIPT_LIBRARY_PATH+"</blockquote></font></blockquote><hr>"
				On Error Goto 0
				Response.End
			ElseIf Session("BrowserAgent") = "On" Then
				Session("BrowserAgent") = "Active"
				'here we can parse the useragent and other variables
				Session("BA_Version") 		= Request.Form("version")
				Session("BA_AppName") 		= Request.Form("appname")
				Session("BA_UserAgent")		= Request.Form("useragent")
				Session("BA_DOM")			= Request.Form("dom")
				Session("BA_NOIFRAME")		= Request.Form("noiframe")
				Session("BA_NOSCRIPT")		= Request.Form("noscript")
				Session("BA_SCREEN")		= Request.Form("screen")
			End If

			mVer	= Session("BA_Version")
			mApp	= Session("BA_AppName")
			mUa 	= LCase(Session("BA_UserAgent"))
			mScr	= Split(Session("BA_SCREEN"),"x")
						
			'set screen resolution
			If UBound(mScr)=2 Then
				ScreenWidth 	= CInt(mScr(0))
				ScreenHeight	= CInt(mScr(1))
				ScreenColor		= CInt(mScr(2))
			End If
			
			HasNoIFRAME	= IIf(LCase(Session("BA_NOIFRAME"))="true",True,False)
			HasNoScript	= IIf(LCase(Session("BA_NOSCRIPT"))="true",True,False)
			
			IsNoScript 	= HasNoScript 'deprecated: only for Backward compatibility
			
			IsDOM 		= IIf(Session("BA_DOM")="true",True,False)

			'always check for safari & opera 
			IsSafari = InStr(1,mUa,"safari")>0		
			IsOpera = InStr(1,mUa,"opera")>0	'before ns or ie

			IsNS = (Not IsOpera) AND (Not IsSafari) AND (mApp="Netscape" Or InStr(1,mUa,"netscape")>0)
			IsMSIE = (Not IsOpera) AND (InStr(1,mUa,"msie")>0)
			IsGecko = InStr(1,mUa,"gecko")>0 ' check for gecko engine

			IsWin32	= InStr(1,mUa,"win")>0
			IsMac 	= InStr(1,mUa,"mac")>0
			IsUnix	= InStr(1,mUa,"unix")>0

			If IsNumeric(mVer) Then mVer = CInt(mVer)
			If IsMSIE Then
				mVer = IIF(InStr(1,mUa,"msie 4")>0,4,mVer)
				mVer = IIF(InStr(1,mUa,"msie 5")>0,5,mVer)
				mVer = IIF(InStr(1,mUa,"msie 5.5")>0,5.5,mVer)
				mVer = IIF(InStr(1,mUa,"msie 6")>0,6,mVer)
				mVer = IIF(InStr(1,mUa,"msie 7")>0,7,mVer)
			ElseIf IsOpera Then
				mVer = CDbl(Mid(mUa,InStr(mUa,"opera")+6,1)) ' set opera version
			ElseIf IsSafari Then
				IsNS = IIf(mVer >= 5,True,False) ' ns6 compatible correct?
			End If	
		End Sub
	
		Public Property Get UserAgent()
			UserAgent = mUa
		End Property
		
		Public Property Get Version()
			Version = mVer
		End Property
		
		Public Property Get AppName()
			AppName = mApp
		End Property
		
	End Class
	
%>