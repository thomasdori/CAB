<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_ImageList.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server ImageList Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">ImageList<hr></Span>
<br>By Raymond Irving (xwisdom at GotDotNet)<br>
<%Page.OpenForm%>
	<%imgList%><BR>
	<%=imgList.GetImageTag("book",100,100,"Cool Book")%>
	<%=imgList.GetImageTag("book",100,100,"Cool Book")%>
	<%=imgList.GetImageTag("clear_all",50,60,"Cool Header")%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim imgList
	
	Public Function Page_Init()
		Set imgList = New_ServerImageList("imgList")
		
		imgList.AddImage "book","images/book01.gif"
		imgList.AddImage "book1","images/book.gif"
		imgList.AddImage "clear_all","images/clear_all.gif"
		imgList.AddImage "header","images/header.gif"
		
	End Function

	Public Function Page_Controls_Init()						
	
	End Function
	
	Public Function Page_Load()
	End Function
		
	Public Function Page_LoadViewState()
		intCounter = CInt(Page.ViewState.GetValue("Count"))
	End Function
%>