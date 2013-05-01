<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DataGrid Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<STYLE>
		.THEADER { background-color:#DDDDDD; }
		.TSELECTEDITEM { background-color:#AAAAAA;color:red; }
		.TALTITEM  {background-color:#cccc99}
		.THEADESTYLE  { font-weight:bold;color:white;background-color:#777777; }

</STYLE>
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%	Page.Execute %>
<BR>
<Span Class="Caption">DataGrid Example</Span>
<span><br>Check the code behing and the properties of the ServerDataGrid and the Pager (ServerDataPager). You can change ANYTHING in the look and feel and behavior of the datagrid...
<br>In the page I also commented out a query that returns 800+ rows. You can use it to test the render peformace. Is never good to render that many rows... check how fast it is when you
enable pagination vs not doing it...

<%Page.OpenForm%>
<%chkAllowPaging%> | <%chkPagerStyle%>
<HR>
<%objDataGrid%>	
<HR>
<%cmdShowDebug%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim cmdShowDebug
	Dim objDataGrid
	Dim chkAllowPaging
	Dim chkPagerStyle
	Dim SortOrder
	
	Dim txtEditContact
	Dim txtEditAddress
	Dim cboEditRegion
	Dim chkEditChk
		
	Public Function Page_Init()
		Set cmdShowDebug = New_ServerLinkButton("cmdShowDebug")		
		Set chkAllowPaging = New_ServerCheckBox("chkAllowPaging")
		Set objDataGrid = New_ServerDataGrid("objDataGrid")
		Set chkPagerStyle  = New_ServerCheckBox("chkPagerStyle")
		
		Set txtEditContact = New_ServerTextBoxEx("txtEditContact",25,30)
		Set txtEditAddress = New_ServerTextBoxEx("txtEditAddress",25,20)
		Set cboEditRegion  = New_ServerDropDown("cboEditRegion")
		Set chkEditChk	   = New_ServerCheckBox("chkEditChk")
		Page.AutoResetScrollPosition = True		
	End Function

	Public Function Page_Controls_Init()						
		SortOrder = "CompanyName"
		cmdShowDebug.Text = "Post..."
		chkAllowPaging.Caption = "Allow Pagination"
		chkPagerStyle.Caption  = "Multi-Page Pager"
		chkAllowPaging.AutoPostBack=True
		chkPagerStyle.AutoPostBack=True
		objDataGrid.Pager.PagerSize = 5
		objDataGrid.Pager.CurrentPageStyle = "color:blue;font-weight:bold"		
		cboEditRegion.Bind GetRecordSet("SELECT rtrim(RegionDescription) as RegionDescription From Region"),"RegionDescription","RegionDescription","",True
		objDataGrid.AlternatingItemCssClass = "TALTITEM"
		objDataGrid.TableWidth = "100%"
		objDataGrid.SortStyle = "color:white"
		objDataGrid.ItemStyle ="background-color:white"
		objDataGrid.SelectedItemStyle = "background-color:#666633;color:white"
		objDataGrid.EditItemStyle = "background-color:#444433;color:white"
		objDataGrid.Control.Style = "border-color:#666633;border-collapse:collapse"
		objDataGrid.HeaderStyle = "font-weight:bold;color:white;background-color:Olive;height:30px"
		objDataGrid.BorderWidth = 1
		objDataGrid.AutoGenerateColumns = False 'To avoid the grid control from doing this :-)

	End Function
	
	Public Function Page_LoadViewState()
		SortOrder  = Page.ViewState.GetValue("SortOrder")	
	End Function

	Public Function Page_PreRender()
		DoSearch
		Page.ViewState.Add "SortOrder",SortOrder
	End Function
	
	Public Sub DoSearch()		
		'Three ways. I recommend the 3rd as it is more flexble...
		'Dim columns(x) and then do Set columns(x) = New_DataGridColumn(ColumnType,DataTextField,DataValueField,DataFormatString)
		'objDataGrid.GenerateColumns() 'Do it and I will take over...
		objDataGrid.CreateColumns Array("CustomerID","CompanyName","Contact","Address","Region","Chk",""),Array("CustomerID","Company","Contact","Address","Region","Chk","Actions")
		'Function to render the header

		Set objDataGrid.DataSource = GetRecordSet("Select CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address,Region,1 as Chk From Customers Order By " & SortOrder)
		objDataGrid.HeaderFunctionName = "objDataGrid_Header"
		'objDataGrid.FooterFunctionName = "objDataGrid_Footer"
		
		'Capture creation of a row...
		objDataGrid.RowDataBound = "objDataGrid_RowDataBound"
		
		objDataGrid.PagerStyle = "background-color:yellow"
		
		objDataGrid.Columns(0).ColumnType = 3   'Templated Column
		objDataGrid.Columns(0).CellRenderFunctionName = "RenderColumn0"
		
		objDataGrid.Columns(1).HeaderText = SortOrder	
		objDataGrid.Columns(1).SortColumn = SortOrder	

		objDataGrid.Columns(4).DataFormatFunctionName = "RenderFormatedCell"
		objDataGrid.Columns(4).HorizonalAlign = "center"
		
		Set objDataGrid.Columns(2).EditControl = txtEditContact
		Set objDataGrid.Columns(3).EditControl = txtEditAddress
		Set objDataGrid.Columns(4).EditControl = cboEditRegion
		Set objDataGrid.Columns(5).EditControl = chkEditChk
		objDataGrid.DataKeyFieldName = "CustomerID"
		objDataGrid.Columns(6).NoWrap = True		
		objDataGrid.Columns(6).ColumnType = 6
		objDataGrid.Columns(6).Style = "background-color:#cccc99;border:dotted 1px black"		
		
		'objDataGrid.Columns(0).Locked = True
		'objDataGrid.Columns(1).Locked = True
	End Sub
	
	Public Function objDataGrid_Edit(e)
		objDataGrid.EditItemIndex = Cint(e.Instance)
	End Function

	Public Function objDataGrid_Cancel(e)
		objDataGrid.EditItemIndex = -1
	End Function

	Public Function objDataGrid_Update(e)
		objDataGrid.EditItemIndex = -1		
		Response.Write "Colums Updated: (Key=" & e.Instance & ") " + txtEditContact.Text & ":" & txtEditAddress.Text & ":" & cboEditRegion.Value & ":" & chkEditChk.Checked
	End Function
	
	Public Function chkAllowPaging_Click()
		objDataGrid.AllowPaging = (chkAllowPaging.Checked)
	End Function
	
	Public Function chkPagerStyle_Click()
		objDataGrid.Pager.PagerType = IIF(chkPagerStyle.Checked,1,0)
	End Function	

	Public Function objDataGrid_ClickColumn0(e)
		objDataGrid.SelectedItemIndex = CInt(e.Instance)
	End Function
	
	Public Function RenderColumn0(ds)
		 Response.Write " <A style='color:black' " 
		 Response.Write Page.GetEventScript("href","objDataGrid","ClickColumn0",ds.AbsolutePosition,"chris")
		 Response.Write " >" & ds(0) & "</a>"
	End Function
	
	public Function objDataGrid_ColumnSort(e)
		If e.Instance = "CompanyName" Then
			SortOrder  = "CompanyName DESC"
		Else
			SortOrder  = "CompanyName"
		End If
		objDataGrid.SelectedItemIndex = -1
		objDataGrid.EditItemIndex= -1
	End Function
		
	Public Function RenderFormatedCell(v)
		If "" & v = "" Then
			RenderFormatedCell =  "@@"
		Else
			RenderFormatedCell =  v
		End If
	End Function
	
	Public Function objDataGrid_RowDataBound(RowContext) 
		If left(RowContext.DataSource.Fields(0),1)  = "B" Then
			If objDataGrid.EditItemIndex = RowContext.DataSource.AbsolutePosition Then
				RowContext.Style= "background-color:#AAAAAA;color:white"
				RowContext.DataSource.Fields(1).Value = "CLASP: Editing?"
			Else
				RowContext.Style= "background-color:#ffcc00;color:white"
				RowContext.DataSource.Fields(1).Value = "CLASP modified this at render time!"
			End If
		End If		
	End Function
	
	Public Function objDataGrid_Header()
		Response.Write "THIS IS THE TABLE HEADER...."
	End Function
	
	
%>