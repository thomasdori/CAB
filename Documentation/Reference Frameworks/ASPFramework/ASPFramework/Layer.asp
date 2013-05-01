<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Layer.asp"    -->
<!--#Include File = "Controls\Server_Panel.asp"     -->
<!--#Include File = "Controls\Server_Graph.asp" -->
<!--#Include File = "Controls\Server_Image.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Layer Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Page.Execute()%>	
<Span Class="Caption">Layer<hr></Span>
<Span>Drag the layers about then click on the "Increment Counter" Link</Span>
<%Page.OpenForm%>
	<%lyrCounter%>
	<%lyrDrag%>
	<%lyrDrag2%>
	<%cmdIncrement%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lyrCounter
	Dim lyrDrag
	Dim lyrDrag2
	Dim cmdIncrement
	Dim panel
	Dim gphData
	Dim imgBook

	Dim intCounter
	
	Public Function Page_Init()
		Set lyrCounter = New_ServerLayer("lyrCounter","Counter",100,250,100,60,"yellow")
		Set lyrDrag = New_ServerLayer("lyrDrag","",250,250,204,154,"limegreen")
		Set lyrDrag2 = New_ServerLayer("lyrDrag2","<center>Drag Here</center>",500,250,130,153,"aqua")
		Set cmdIncrement  = New_ServerLinkButton("cmdIncrement")
		Set panel = New_ServerPanel("panel",200,150)
		Set gphData = New_ServerGraph("gphData")
		Set imgBook = New_ServerImage("imgBook")		
		
		lyrDrag.AddContent  panel		' add the panel control to lyrDrag layer
		lyrDrag2.AddContent imgBook	' add the panel control to lyrDrag layer
		
		'add some additional text to lyrCounter
		lyrCounter.AddContent  "<hr size='1' width='100'>"

	End Function

	Public Function Page_Controls_Init()						
		intCounter = 0
		cmdIncrement.Text = "Increment<br>counter"
		imgBook.ImageSrc = "Images/book01.gif"

		lyrDrag.DragEnabled = True
		lyrDrag.RaiseOnMove = True
		
		lyrDrag2.DragEnabled = True
		lyrDrag2.Control.Style = "border:1px solid black;"
		
		lyrCounter.AutoSize = True 'auto size layer
		lyrCounter.Control.Style = "border:1px solid silver;"			
	
		panel.Mode = 2 
		panel.PanelTemplate = "RenderPanelContent"

		gphData.Mode 	= 1
		gphData.Width	= 140
		gphData.Height	= 100

		Page.ViewState.Add "Count",intCounter		
	End Function

	Function Page_Load()
		gphData.XScale = 1
		gphData.XScaleTick = 1
		gphData.YScale = 5000
		
	
		' add row data
		gphData.AddRow(Array(4208, 10535, 10762, 10517, 4066, 4477, 10547))
		gphData.AddRow(Array(2178, 4017, 3803, 22481, 2733, 2385, 14072))
		gphData.AddRow(Array(1000, 2000, 3000, 400, 1000, 2000, 300))

	End Function
	
	Public Function Page_LoadViewState()
		intCounter = CInt(Page.ViewState.GetValue("Count"))
	End Function
	
	Public Function cmdIncrement_OnClick()
		lyrCounter.InnerHTML = "Count:" & intCounter
		intCounter = intCounter + 1
		Page.ViewState.Add "Count",intCounter
	End Function

	Public Function RenderPanelContent()
		Response.Write "This a 3D Panel inside a drag-able <b>Layer</b>."
		Response.Write "<center>"
		Call gphData.Render()
		Response.Write "</center>"
	End Function

	Public Function lyrDrag_OnMove(obj,args)
		lyrCounter.AddContent "<span>Layer Moved to location "& lyrDrag.Left &","& lyrDrag.Top &"</span><br>&nbsp;"
	End Function
%>