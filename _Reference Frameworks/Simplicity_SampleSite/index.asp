<%

OPTION EXPLICIT

' ---------------------------------------------------------------
'	Appllication Constants
' ---------------------------------------------------------------
	
	dim VERSION
	dim EXT,PAGEEXT, ALLOWCHARS
	dim EMAILERRORS,ERRORCHECK
	dim CACHECONTROL,SHOWPROFILER
	dim write
	
	VERSION			= "0.6"
	write 			= ""
	
	EXT 			= ".asp"
	PAGEEXT 		= ".html"
	ALLOWCHARS 		= "[^a-z0-9\-_:&=.?/$]"	' RegEx for allowable URL characters
	
	ERRORCHECK		= false	' true = YES / false = NO
	EMAILERRORS		= false	' true = YES / false = NO
	
	CACHECONTROL	= true	' true = ON / false = OFF
	
	SHOWPROFILER	= false	' true = ON / false = OFF
	
' ---------------------------------------------------------------
'  ASP errors On
' ---------------------------------------------------------------
	
	if ERRORCHECK then
			On Error Resume NEXT 
	end if
	
' ---------------------------------------------------------------
'  Start ASP Timer
' ---------------------------------------------------------------
	
	dim STARTTIME
	STARTTIME = Timer()
	
' ---------------------------------------------------------------
'  Load FrameWork
' ---------------------------------------------------------------
%>
<!-- #include virtual = "/system/core/aspFramework.asp" -->