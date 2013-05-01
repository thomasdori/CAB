<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_Tab.asp"    -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Tabs</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">TabStrip Control Sample...<HR> </Span>
<Span><br>TABS Is in Beta!!!<BR>There is more room for improvements, but by know it supports many features, including AutoPostback</Span>
<%Page.OpenForm%>
<%MainTab%>
<HR>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim MainTab
	Dim SubTab
	Dim lnk 
	Dim txtName
	Dim objDataGrid
	
	
	
	Public Function Page_Init()
		Dim x
		Set txtName = New_ServerTextBox("txtName")
		Set lnk  = New_ServerLinkButton("lnk")		
		
		Set objDataGrid = New_ServerDataGrid("objDataGrid")
			DataGrid_BlueTemplate objDataGrid,True
			objDataGrid.AllowPaging = True
		
		Set MainTab  = New_ServerTabStrip("MainTab")				
		SetTabStripStyle MainTab
		
		Dim Tabs(7)
		For x=0 To Ubound(Tabs)
			Set Tabs(x) = New ServerTab
			SetTabStyle Tabs(x),"Tab #" & x,"TabX"			
		Next
		Tabs(0).RenderFunction  = "Tab00"
		Tabs(0).Caption  = "General"
		
		Tabs(1).Caption  = "Security"		
		Tabs(1).RenderFunction  = "TabSecurity"		
		
		Tabs(2).Caption  = "Users"		
		Tabs(2).RenderFunction  = "Tab02"

		MainTab.Tabs = Tabs
		
		lnk.Text = "Next..."
	End Function

	Public Function Page_Controls_Init()		
	End Function
	
	Public Function Page_Load()

	End Function

	Public Function SetTabStripStyle(oTabStrip)
			oTabStrip.AutoPostBack = False
			oTabStrip.TabStripTabSeparatorTemplate = "TabSepFnc"
			oTabStrip.TabsPerRow = 10
			oTabStrip.TabsWidth = "80%"
			oTabStrip.Width = "100%"
			oTabStrip.Height = "300"
			oTabStrip.BorderWidth = 1
			oTabStrip.BorderColor = "#da9e86"
			oTabStrip.Style = "background-color:#fffcea;"
	End Function
	
	Public Function SetTabStyle(oTab,Caption,fncRender)
			oTab.Caption = Caption
			oTab.Style = "background-color:#F6D7CE;BORDER-TOP: #da9e86 1px solid;BORDER-RIGHT: #da9e86 1px solid"
			oTab.RenderFunction = "TabX"
			oTab.SelectedStyle  = "background-color:#F6D7CE;BORDER-TOP: #da9e86 1px solid;BORDER-RIGHT: #da9e86 1px solid;background-color:#EDB5A6;"
			oTab.HoverStyle  = "background-color:#F6D7CE;BORDER-TOP: #da9e86 1px solid;BORDER-RIGHT: #da9e86 1px solid;color:red;text-decoration:underline;"
			oTab.SelectedHoverStyle	= "background-color:#F6D7CE;BORDER-TOP: #da9e86 1px solid;BORDER-RIGHT: #da9e86 1px solid;background-color:#EDB5A6;color:red;"
			oTab.RenderFunction = fncRender
	End Function

	Public Function lnk_OnClick()
		MainTab.SelectedTabIndex= MainTab.SelectedTabIndex + 1
	End Function
	
	
	Public Function TabSepFnc() 
			Response.Write "<div style='width:80%;height:5px;background-color:#EDB5A6;font-size:1px'></div>"
	End Function
%>


<%Public Function TabX()
	Dim x,mx
	Randomize Timer
	mx = int(rnd*30) + 1
	for x=1 to mx
		Response.Write "HERE" & x & "<BR>"	
	next
End Function%>


<%Public Function Tab00()%>
	<Table><TR><TD>First Name</TD><TD><%txtName%></TD></TR></TABLE>
<%End Function%>

<%Public Function Tab02()%>
	This is a good example of a division with all sort of things...<BR>
	To Move Next Tab Click Here <%lnk%>
<%End Function%>

<%Public Function TabSecurity()
	Set objDataGrid.DataSource = GetRecordSet("Select CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address From Customers")
	objDataGrid.Control.Style = "border-collapse:collapse;width:100%"
	objDataGrid.Render()
End Function%>

