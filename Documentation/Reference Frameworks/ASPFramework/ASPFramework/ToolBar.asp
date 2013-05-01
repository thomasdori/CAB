<!--#Include File = "Controls\WebControl.asp"       -->
<!--#Include File = "Controls\Server_CheckBox.asp"     -->
<!--#Include File = "Controls\Server_ToolBar.asp"     -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>ToolBar Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<!--#Include File = "Home.asp"        -->
<style>
	.imgBtn {
		border: 1px solid #efedde;
	}
	.link {
		width:75px;
		text-align:center;
		border:1px solid silver;
		color:black;
		background-color:#eeeeee;
		padding: 3px;
		cursor:default;
		background-image:url('images/bgstrip.gif');
	}
	.linkHover {
		width:75px;
		text-align:center;
		color:black;
		background-color:#abbee8;
		border:1px solid navy;
		padding: 3px;		
		cursor:default;
		background-image:url('images/bgstrip2.gif');
	}
</style>
</HEAD>
<BODY>
<%Call Page.Execute()%>	
<Span Class="Caption">TOOLBAR EXAMPLE</Span>

<%Call Page.OpenForm()%>
<table border="0" cellpadding="4" cellspacing="2" width="100%">
  <tr>
    <td width="11%" valign="top"><%tBar2%></td>
    <td width="89%" valign="top">
      <table border="0" width="100%" cellspacing="2" cellpadding="1">
        <tr>
          <td width="100%"><%tBar3%></td>
        </tr>
        <tr>
          <td width="100%"><%tBar%></td>
        </tr>
        <tr>
          <td width="100%"><%tBar1%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<%Call Page.CloseForm()%>
</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	
	Dim tBar
	Dim tBar1
	Dim tBar2
	Dim tBar3
	Dim chkBox
	
	Public Function Page_Init()
		'Page.DebugEnabled = True
		Set tBar = New_ServerToolBar("tBar")
		Set tBar1 = New_ServerToolBar("tBar1")
		Set tBar2 = New_ServerToolBar("tBar2")
		Set tBar3 = New_ServerToolBar("tBar3")
		Set chkBox = New_ServerCheckBox("chkBox")
	End Function
	
	Public Function Page_Controls_Init()

		chkBox.Caption = "Layout&nbsp;"
		chkBox.Checked = True
		chkBox.AutoPostBack = True
		
		tBar.Layout 		= tbHorizontal 
		tBar.ButtonStyle	= "padding:3px"
		tBar.AddButton "nw","New",75,"New"
		tBar.AddButton "sv","Save",75,"Save"
		tBar.AddButton "op","Open",75,"Open"
		tBar.AddSeparator
		tBar.AddButton "bd","Bold",75,"Bold"
		tBar.AddSeparator
		tBar.AddButtonTemplate "RenderCheckBox",0,"Checkbox Control - Change Layout"
		tBar.AddSeparator
		tBar.AddButtonTemplate "RenderLink",0,"URL--"
	
		tBar1.Spacing 				= 2
		tBar1.BackColor 			= "#FFFFFF"
		tBar1.ButtonCssClass		= "link"
		tBar1.ButtonHoverCssClass	= "linkHover"
		tBar1.ButtonSelectCssClass	= "linkHover"
		tBar1.Layout = tbHorizontal 
		tBar1.AddButton "nw","New",75,"New"
		tBar1.AddButton "sv","Save",75,"Save"
		tBar1.AddButton "op","Open",75,"Open"
		tBar1.AddButton "bd","Bold",75,"Bold"
	
		tBar2.Layout 		= tbVertical 
		tBar2.Control.Style = "border:1px outset"
		tBar2.ButtonStyle	= "padding:3px"
		tBar2.AddButton "nw","New",75,"New"
		tBar2.AddSeparator
		tBar2.AddButton "sv","Save",75,"Save"
		tBar2.AddButton "op","Open",75,"Open"
		tBar2.AddButton "bd","Bold",75,"Bold"

		tBar3.Spacing = 2
		tBar3.Layout = tbHorizontal 
		tBar3.ButtonCssClass	= "imgBtn"
		tBar3.Control.Style = "border:1px outset"
		tBar3.AddButton "nw","<img src='"+SCRIPT_LIBRARY_PATH+"richtext/images/preview.gif' hspace='1' vspace='1'>",0,"Preview"
		tBar3.AddSeparator
		tBar3.AddButton "sv","<img src='"+SCRIPT_LIBRARY_PATH+"richtext/images/Save.gif' hspace='1' vspace='1'>",0,"Save"
		tBar3.AddButton "op","<img src='"+SCRIPT_LIBRARY_PATH+"richtext/images/Copy.gif' hspace='1' vspace='1'>",0,"Copy"
		tBar3.AddButton "bd","<img src='"+SCRIPT_LIBRARY_PATH+"richtext/images/Bold.gif' hspace='1' vspace='1'>",0,"Bold"
		
	End Function

	Public Function chkBox_Click()
		tBar.Layout = IIf(chkBox.Checked,tbHorizontal,tbVertical)
	End Function
	
	Public Function tBar2_OnClick(e)
		Dim id
		Dim newState		
		id 			=  e.Instance
		newState	= Not tBar2.GetButtonToggleState(id)		
		tBar2.ToggleAllButton(False)	'reset all selected button
		Call tBar2.ToggleButton(id,newState)
	End Function

	Public Function tBar1_OnClick(e)
		Dim id
		Dim newState		
		id 			=  e.Instance
		newState	= Not tBar1.GetButtonToggleState(id)		
		tBar1.ToggleAllButton(False)	'reset all selected button
		Call tBar1.ToggleButton(id,newState)
	End Function
	
	Public Function RenderCheckBox()
		chkBox.Render()
	End Function
	
	Public Function RenderLink() 
		Response.Write "<a href='http:\\www.yahoo.com' target='_blank'>Yahoo</a>"
	End Function
%>