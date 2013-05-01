<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_HyperLink.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>HyperLink Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">HyperLink Example<hr></Span>

<%Page.OpenForm%>
	<%lnk01%><BR>
	<%lnk02%><BR>
	<%lnk03%><BR>
	
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lnk01,lnk02,lnk03
	
	Public Function Page_Init()
		Set lnk01 = New_ServerHyperLink("lnk01")
		Set lnk02 = New_ServerHyperLink("lnk02")
		Set lnk03 = New_ServerHyperLink("lnk03")
	End Function

	Public Function Page_Controls_Init()						
		 lnk01.Text  ="Link to another page"
		 lnk01.NavigateURL = "DataGrid.asp"
		 
		 lnk02.Text  = "Opens in new window"
		 lnk02.Target = "_blank"
		 lnk02.NavigateURL = "DataGrid.asp"
		 
		 lnk03.ImageURL = "images/book01.gif"
		 lnk03.NavigateURL = "DataGrid.asp"
	End Function
%>