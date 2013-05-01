<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_ProgressBar.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>PlaceHolder Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">ProgressBar Example</Span>
<%Page.OpenForm%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	
	Public Function Page_Init()
	Response.Write "<script language='javascript'>CLASP_ProgressBar_Start();</script>"
'	Response.Flush
	Dim x,i
	For x = 0 to 100 step 5
		Response.Write "<script language='javascript'>CLASP_ProgressBar_Update(" & x & ");</script>"
		Response.Flush
		For i=1 to 500000
		next
	Next
	Response.Write "<script language='javascript'>CLASP_ProgressBar_End();</script>"

	
	End Function

	Public Function Page_Load()
		
	End Function
	
%>