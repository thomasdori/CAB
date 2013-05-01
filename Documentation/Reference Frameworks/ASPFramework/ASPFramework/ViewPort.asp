<!--#Include File = "Controls\WebControl.asp"       -->
<!--#Include File = "Controls\Server_Button.asp"     -->
<!--#Include File = "Controls\Server_HyperLink.asp"     -->
<!--#Include File = "Controls\Server_CheckBox.asp"     -->
<!--#Include File = "Controls\Server_ViewPort.asp"     -->
<!--#Include File = "Controls\Server_ToolBar.asp"     -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>ViewPort Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<!--#Include File = "Home.asp"        -->
<style>
	.linkSel {
		width:75px;
		text-align:center;
		text-decoration:none;
		vertical-align:middle;
		border:1px solid navy;
		color:white;
		background-color:navy;
		padding: 3px;
	}
	.link {
		width:75px;
		text-align:center;
		text-decoration:none;
		vertical-align:middle;
		border:1px solid silver;
		color:black;
		background-color:#eeeeee;
		padding: 3px;
	}
	.link:hover {
		color:black;
		background-color:#abbee8;
		border:1px solid navy;
	}
</style>
</HEAD>
<BODY>
<%Call Page.Execute()%>	
<Span Class="Caption">VIEWPORT EXAMPLE</Span>

<%Call Page.OpenForm()%>
<table border="0" cellpadding="4" width="100%">
  <tr>
    <td width="11%" valign="top"><%pback%></td>
    <td width="89%" valign="top">Last ViewPort URL Since PostBack: <%=vport.LocationURL%></td>
  </tr>
  <tr>
    <td width="11%" valign="top"><%link1%><br><%link2%></td>
    <td width="89%" valign="top">
      <table border="0" width="100%">
        <tr>
          <td width="100%" bgcolor="#FAF5D8"><%chkHideShow%> <%chkStretch%> <%chkSBar%> <%chkAHeight%></td>
        </tr>
        <tr>
          <td width="100%"><%vport%></td>
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
	Dim vport
	Dim link1
	Dim link2
	Dim pback
	Dim chkHideShow
	Dim chkSBar
	Dim chkStretch
	Dim chkAHeight	

	Public Function Page_Init()
		'Page.DebugEnabled = True
		Set vport = New_ServerViewPort("vport","",200,100)
		Set link1 = New_ServerHyperLink("link1")
		Set link2 = New_ServerHyperLink("link2")
		Set pback = New_ServerLinkButton("pback")
		Set chkHideShow = New_ServerCheckBox("chkHideShow")
		Set chkSBar = New_ServerCheckBox("chkSBar")
		Set chkStretch = New_ServerCheckBox("chkStretch")
		Set chkAHeight = New_ServerCheckBox("chkAHeight")
				
		Page.PageID = "Main"			' make this page and viewstate unique when using server-side or custom viewstate
		'Page.PreserveViewState = True	' preserve the content of the viewstate when navigating across pages. PageID must be set
		
	End Function
	
	Public Function Page_Controls_Init()						
		link1.Control.CssClass = "link"
		link1.Text = "Grid"
		link1.Target = "vport"
		link1.NavigateURL = "datagrid.asp"

		link2.Control.CssClass = "link"
		link2.Text = "Drill-Down"
		link2.Target = "vport"
		link2.NavigateURL = "drilldown.asp"

		pback.Control.CssClass = "link"
		pback.Text = "PostBack"

		chkHideShow.Caption = "Hide/Show ViewPort"
		chkHideShow.AutoPostBack = True
		chkHideShow.Checked = True

		chkSBar.Caption = "Hide/Show ScrollBar"
		chkSBar.AutoPostBack = True
		chkSBar.Checked = True

		chkStretch.Caption = "Stretch To Fit"
		chkStretch.AutoPostBack = True
		chkStretch.Checked = True
		
		chkAHeight.Caption = "Auto Adjust Height"
		chkAHeight.AutoPostBack = True

		vport.StretchToFit = True
		vPort.InnerHTML ="<h2>The ViewPort</h2>"
	End Function

	Public Function pback_OnClick()
		'pback.Control.CssClass = "linkSel"
	End Function
	
	Public Function chkHideShow_Click()
		vport.Control.Visible = chkHideShow.Checked		
	End Function

	Public Function chkSBar_Click()
		vport.ScrollBarMode = IIf(chkSBar.Checked,1,0)
	End Function

	Public Function chkStretch_Click()
		vport.StretchToFit = chkStretch.Checked		
	End Function

	Public Function chkAHeight_Click()
		vport.AutoHeight = chkAHeight.Checked		
	End Function
%>