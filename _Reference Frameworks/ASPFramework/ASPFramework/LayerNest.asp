<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Layer.asp"    -->
<!--#Include File = "Controls\Server_Image.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Nested Layer Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%Page.Execute()%>	
<Span Class="Caption">NESTED LAYER EXAMPLE</Span>
<table><tr><td>In this example you can drag the blue layer and click on layers 1 to 3</td></tr></table>
<%Page.OpenForm%>
	<%lyrBase%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...

	Dim lyrBase
	Dim lyrText
	Dim lyrBtn1
	Dim lyrBtn2
	Dim lyrBtn3
	Dim imgBook
	
	
	Public Function Page_Init()
		Set lyrBase = New_ServerLayer("lyrBase","<font color='#ffffff' size='2'>&nbsp;Title:Drag Here</font>",530,220,200,200,"navy")
		Set lyrText = New_ServerLayer("lyrText","<br><center><h2>Click on layers 1 to 3</h2>",10,20,180,135,"limegreen")
		Set lyrBtn1 = New_ServerLayer("lyrBtn1","1",15,170,50,25,"pink")
		Set lyrBtn2 = New_ServerLayer("lyrBtn2","2",70,170,50,25,"red")
		Set lyrBtn3 = New_ServerLayer("lyrBtn3","3",125,170,60,25,"yellow")

		Set imgBook = New_ServerImage("imgBook")		
		
		lyrBase.AddContent lyrText	' add the text layer the base layer		
		lyrBase.AddContent lyrBtn1	' add the button layer the base layer		
		lyrBase.AddContent lyrBtn2	' add the button layer the base layer		
		lyrBase.AddContent lyrBtn3	' add the button layer the base layer		

		
		lyrText.AddContent imgBook

	End Function

	Public Function Page_Controls_Init()						
		lyrBase.DragEnabled = True	
		'lyrText.DragEnabled = True	
		lyrBtn1.AutoPostBack = True
		lyrBtn2.AutoPostBack = True
		lyrBtn3.AutoPostBack = True

		lyrBtn1.PreventBubble = True
		lyrBtn2.PreventBubble = True
		lyrBtn3.PreventBubble = True
		lyrText.PreventBubble = True

		imgBook.ImageSrc = "images/clear_all.gif"
	End Function

	Public Function Page_Load()
		Page.RegisterEventListener lyrBtn1,"onclick","window.open(""datagrid.asp"")"
	End Function

	Public Function lyrBtn1_Click()
		lyrText.InnerHTML = "<br><center><h2>Hello From Button 1</h2>"
		imgBook.ImageSrc = "images/clear_all.gif"
	End Function

	Public Function lyrBtn2_Click()
		lyrText.InnerHTML = "<br><center><h2>Hello From <u>Button</u> 2</h2>"
		imgBook.ImageSrc = "images/book.gif"
	End Function

	Public Function lyrBtn3_Click()
		lyrText.InnerHTML = "<br><center><h2>Hello From <i>Button</i> 3</h2>"
		imgBook.ImageSrc = "images/clear_all.gif"
	End Function
	
%>