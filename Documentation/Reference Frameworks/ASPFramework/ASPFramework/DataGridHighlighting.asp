<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DataGrid Row Highlight Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%Page.Execute%>
<BR>
<Span Class="Caption">DataGrid Row Highlight Example</Span>
<%Page.OpenForm%>
<%chkAllowPaging%> | <%chkPagerStyle%> | <%chkLight%>
<HR>
<%objDataGrid%>	
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim objDataGrid
	Dim chkAllowPaging
	Dim chkPagerStyle
	Dim chkLight
			
	Public Function Page_Init()
		Set objDataGrid = New_ServerDataGrid("objDataGrid")
		Set chkAllowPaging = New_ServerCheckBox("chkAllowPaging")
		Set chkPagerStyle  = New_ServerCheckBox("chkPagerStyle")
		Set chkLight = New_ServerCheckBox("chkLight")

		Page.AutoResetScrollPosition = True		
	End Function

	Public Function Page_Controls_Init()						
		objDataGrid.CursorHighlighting = True
		objDataGrid.Pager.PagerSize = 5
		objDataGrid.Pager.CurrentPageStyle = "color:blue;font-weight:bold"		

		chkAllowPaging.Caption = "Allow Pagination"
		chkAllowPaging.AutoPostBack=True

		chkPagerStyle.Caption  = "Multi-Page Pager"
		chkPagerStyle.AutoPostBack=True

		chkLight.Caption  = "Row Highlighting"
		chkLight.AutoPostBack=True
		chkLight.Checked = True

	End Function

	Public Function Page_PreRender()
		DoSearch
	End Function

	Public Function chkAllowPaging_Click()
		objDataGrid.AllowPaging = (chkAllowPaging.Checked)
	End Function
	
	Public Function chkPagerStyle_Click()
		objDataGrid.Pager.PagerType = IIF(chkPagerStyle.Checked,1,0)
	End Function

	Public Function chkLight_Click()
		objDataGrid.CursorHighlighting = IIF(chkLight.Checked,1,0)
	End Function
	
	Public Sub DoSearch()		
		Set objDataGrid.DataSource = GetRecordSet("Select CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address,Region From Customers Order By CompanyName")
		objDataGrid.AlternatingItemStyle = "background-color:beige;"
		objDataGrid.TableWidth = "100%"
		objDataGrid.SortStyle = "color:white"
		objDataGrid.SelectedItemStyle = "background-color:#666633;color:white"
		objDataGrid.EditItemStyle = "background-color:#444433;color:white"
		objDataGrid.Control.Style = "cursor:default;border-color:#666633;border-collapse:collapse"
		objDataGrid.HeaderStyle = "font-weight:bold;color:white;background-color:Olive;height:23px"
		objDataGrid.BorderWidth = 1
		objDataGrid.AutoGenerateColumns = False 'To avoid the grid control from doing this :-)
		'Three ways. I recommend the 3rd as it is more flexble...
		'Dim columns(x) and then do Set columns(x) = New_DataGridColumn(ColumnType,DataTextField,DataValueField,DataFormatString)
		'objDataGrid.GenerateColumns() 'Do it and I will take over...
		objDataGrid.CreateColumns Array("CustomerID","CompanyName","Contact","Address","Region"),Array("CustomerID","Company","Contact","Address","Region")

	End Sub
	
	Public Function objDataGrid_OnClick(e)
		objDataGrid.SelectedItemIndex = CInt(e.Instance)		
	End Function
	
%>