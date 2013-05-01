<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_Graph.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Graph Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">

</HEAD>
<BODY id="CLASPBody">
<!-- Include File = "Home.asp" -->
<%Page.Execute%>	
<Span Class="Caption">Graph<hr></Span>
<%Page.OpenForm%>
<%cboGraphs%> | <%chkLegend%>
<hr>
<%gphData%>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim gphData
	Dim chkLegend
	Dim cboGraphs
	
	Dim mShowLegend
	
'PAGE EVENT HANDLERS	
	Function Page_Init()		

 		Set chkLegend	= New_ServerCheckBox("chkLegend")
		Set cboGraphs	= New_ServerDropDown("cboGraphs")
		Set gphData 		= New_ServerGraph("gphData")

		chkLegend.AutoPostBack = True
		cboGraphs.AutoPostBack = True
		
	End Function

	Function Page_Controls_Init()	
		chkLegend.Caption = "Show Legend"
		chkLegend.Checked = True
		cboGraphs.Items.Append "Simple Bars",1,True
		cboGraphs.Items.Append "Stacked Bars",2,False
		cboGraphs.Items.Append "Stacked Bars (Area)",3,False

		gphData.Mode 		= 1
		gphData.Width		= 450
		gphData.Title 	= "Web Page Hits"
		gphData.XLabel 	= "MONTHS"
		gphData.YLabel 	= "HITS"		
		'gphData.Control.Style = "background-color:yellow"		
	End Function

	Function Page_Load()
		gphData.XScale = 1
		gphData.XScaleTick = 1
		gphData.YScale = 5000
		
	
		' add row data
		gphData.AddRow(Array(4208, 10535, 10762, 10517, 4066, 4477, 10547))
		gphData.AddRow(Array(2178, 4017, 3803, 22481, 2733, 2385, 14072))
		gphData.AddRow(Array(1000, 2000, 3000, 400, 1000, 2000, 300))
		
		If chkLegend.Checked Then 
			gphData.SetLegend(Array("Home Page","Downloads","FAQ"))
		End If

	End Function
	
'WEBCONTROLS POSTBACK EVENT HANDLERS
	Public Function cboGraphs_ItemChange(e)
		gphData.Mode = cboGraphs.Value
	End Function
	

%>