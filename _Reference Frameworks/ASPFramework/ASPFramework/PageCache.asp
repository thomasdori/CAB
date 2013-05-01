<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"    -->
<!--#Include File = "Controls\PageCache.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Page Cache</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<STYLE>
		.THEADER { background-color:#DDDDDD; }
		.TSELECTEDITEM { background-color:#AAAAAA;color:red; }
		.TALTITEM  {background-color:#DDDDDD}
		.THEADESTYLE  { font-weight:bold;color:white;background-color:#777777; }

</STYLE>
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%	Page.Execute %>
<BR>
<Span Class="Caption">Page Caching Sample</Span>
<span><br>This page shows how to cache an ASP page using the new caching "Stuff" :-P<BR>
The Cache is set to expire every 5 minutes and will by-pass postbacks. All this by just adding Controls\PageCache.asp!

<%Page.OpenForm%>
<%chkAllowPaging%> | <%chkPagerStyle%> | <%cmdShowDebug%>
<HR>
<%objDataGrid%>	
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim cmdShowDebug
	Dim objDataGrid
	Dim chkAllowPaging
	Dim chkPagerStyle
	
	Public Function Page_RequestCacheAction(CacheYN)
		'You can access the to PageCache and  configure the cache object further. This is just to know if you want or not to cache
		CacheYN=True 
		PageCache.CacheDuration = 5 'in minutes	
		Response.Write "Version as Of: " & Now & "<BR>" '(this will be written to the cache :-P )
	End Function
	
	Public Function Page_Init()
		
		Set cmdShowDebug = New_ServerLinkButton("cmdShowDebug")
		Set objDataGrid = New ServerDataGrid
			objDataGrid.Control.Name = "objDataGrid"
		Set objDataGrid.DataSource = GetRecordSet("Select ProductID as ID, ProductName as Name, UnitPrice from products")

		DataGrid_RedTemplate objDataGrid,True
		objDataGrid.BorderWidth = 1
		objDataGrid.AutoGenerateColumns = True
		
		Set chkAllowPaging = New_ServerCheckBox("chkAllowPaging")
		Set chkPagerStyle  = New_ServerCheckBox("chkPagerStyle")
		Page.AutoResetScrollPosition = True
	End Function

	Public Function Page_Controls_Init()						
		cmdShowDebug.Text = "Post..."
		chkAllowPaging.Caption = "Allow Pagination"
		chkPagerStyle.Caption  = "Multi-Page Pager"
		chkAllowPaging.AutoPostBack=True
		chkPagerStyle.AutoPostBack=True
		objDataGrid.Pager.PagerSize = 5
		objDataGrid.Pager.CurrentPageStyle = "color:red;font-weight:bold"
	End Function
	
	Public Function chkAllowPaging_Click()
		objDataGrid.AllowPaging = (chkAllowPaging.Checked)
	End Function
	
	Public Function chkPagerStyle_Click()
		objDataGrid.Pager.PagerType = IIF(chkPagerStyle.Checked,1,0)
	End Function	

	
%>