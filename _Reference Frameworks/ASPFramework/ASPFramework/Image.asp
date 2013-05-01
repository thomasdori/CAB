<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Image.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Image Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">Images<hr></Span>
<%Page.OpenForm%>
	<%cmdOK%>
	<%imgBook%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim cmdOK
	Dim imgBook
	
	Public Function Page_Init()
		Set cmdOK = New_ServerLinkButton("cmdGo")		
		Set imgBook = New_ServerImage("imgBook")		
	End Function

	Public Function Page_Controls_Init()						
		cmdOK.Text = "Post back..."
		imgBook.ImageSrc = "Images/book01.gif"
	End Function

%>