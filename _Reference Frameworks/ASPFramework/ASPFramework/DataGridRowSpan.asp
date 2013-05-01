<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DataGrid Row Span Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<style>
.ColumnHeader
{
	font-weight:bold;
	color:white;
	background-color:Olive;
	height:24px;
}
.SortHeader
{
	color:white;
	width:100%;
	padding:2px;
	cursor:default;
}

.SortHeader:hover
{
	color:white;
	text-decoration:none;
	background-color:#909000;
	border-top:1px solid white;
	border-left:1px solid white;
	border-bottom:1px solid black;
	border-right:1px solid black;
}
</style>
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%Page.Execute%>
<BR>
<Span Class="Caption">DataGrid Row Span Example</Span>
<%Page.OpenForm%>
<%chkAllowPaging%> | <%chkPagerStyle%> | <%chkLight%> | <%chkRowSpan%> | <%chkRowCaption%> | <%chkHdrSpan%>
<HR>
<%chkScroll%> | <%chkFix%>
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
	Dim chkRowSpan
	Dim chkRowCaption
	Dim chkHdrSpan
	Dim chkScroll
	Dim chkFix
	
	Dim Section
	Dim SortField
	Dim SortColumn
	Dim SortOrder
			
	Public Function Page_Init()
		Set objDataGrid		= New_ServerDataGrid("objDataGrid")
		Set chkAllowPaging	= New_ServerCheckBox("chkAllowPaging")
		Set chkPagerStyle 	= New_ServerCheckBox("chkPagerStyle")
		Set chkLight		= New_ServerCheckBox("chkLight")
		Set chkRowSpan		= New_ServerCheckBox("chkRowSpan")
		Set chkRowCaption	= New_ServerCheckBox("chkRowCaption")
		Set chkHdrSpan 		= New_ServerCheckBox("chkHdrSpan")
		Set chkScroll		= New_ServerCheckBox("chkScroll")
		Set chkFix			= New_ServerCheckBox("chkFix")

		Page.AutoResetScrollPosition = True		
		Section = 0
		SortColumn = 1
		SortField = "CompanyName"
		
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

		chkRowSpan.Caption  = "Row Span"
		chkRowSpan.AutoPostBack=True
		chkRowSpan.Checked = True

		chkRowCaption.Caption  = "Row Caption"
		chkRowCaption.AutoPostBack=True
		chkRowCaption.Checked = True

		chkHdrSpan.Caption  = "Header Span"
		chkHdrSpan.AutoPostBack=True

		chkScroll.Caption  = "ScrollBars"
		chkScroll.AutoPostBack=True

		chkFix.Caption  = "Fix Column Header"
		chkFix.AutoPostBack=True

	End Function

	Public Function Page_LoadViewState()
		Dim vArr
		SortField = Page.ViewState.GetValue("SortF")
		vArr = Split(Page.ViewState.GetValue("SortC"),".")
		If UBound(vArr)=1 Then
			If IsNumeric(vArr(0)) Then SortColumn = CInt(vArr(0))
			If IsNumeric(vArr(1)) Then SortOrder = CInt(vArr(1))
		End If
	End Function
	
	Public Function Page_SaveViewState()
		Page.ViewState.Add "SortF",SortField
		Page.ViewState.Add "SortC",SortColumn &"." & SortOrder
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

	Public Function chkFix_Click()
		objDataGrid.FixedColumnHeader = IIF(chkFix.Checked,1,0)
	End Function
	
	Public Function chkScroll_Click()
		objDataGrid.TableHeight = 150
		objDataGrid.EnableScrollBars = IIF(chkScroll.Checked,1,0)
	End Function
	
	Public Sub DoSearch()		
		Set objDataGrid.DataSource = GetRecordSet("Select CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address,Region From Customers Order By "+SortField + IIf(SortOrder=0," ASC "," DESC"))
		objDataGrid.AlternatingItemStyle = "background-color:beige;"
		objDataGrid.TableWidth = "100%"
		objDataGrid.SortCssClass = "SortHeader"
		objDataGrid.SelectedItemStyle = "background-color:#666633;color:white"
		objDataGrid.EditItemStyle = "background-color:#444433;color:white"
		objDataGrid.Control.Style = "cursor:default;border-color:#666633;border-collapse:collapse"
		objDataGrid.HeaderCssClass = "ColumnHeader"
		objDataGrid.BorderWidth = 1
		objDataGrid.AutoGenerateColumns = False 'To avoid the grid control from doing this :-)
		
		objDataGrid.CreateColumns Array("CustomerID","CompanyName","Contact","Address","Region"),Array("Customer ID","Company","Contact","Address","Region")
		'NOTE: Three ways. I recommend the 3rd as it is more flexble...
		'Dim columns(x) and then do Set columns(x) = New_DataGridColumn(ColumnType,DataTextField,DataValueField,DataFormatString)
		'objDataGrid.GenerateColumns() 'Do it and I will take over...

		'Capture creation of a row...
		 objDataGrid.RowDataBound = "objDataGrid_RowDataBound"

		If chkRowSpan.Checked Then 
			objDataGrid.Columns(1).RowSpan =3
		End If

		objDataGrid.Columns(0).NoWrap = True
		objDataGrid.Columns(1).SortColumn="1"		
		objDataGrid.Columns(2).SortColumn="2"
		objDataGrid.Columns(3).SortColumn="3"
		
		'setup sort column's sort order
		objDataGrid.Columns(SortColumn).SortOrder = IIf(SortOrder=0,1,2) 'ASC or DESC
		objDataGrid.Columns(SortColumn).Style = "border-left:2px solid olive;border-right:2px solid olive" 'add demarcation to sorted column

		objDataGrid.Columns(1).DataFormatFunctionName = "objDataGrid_ColumnFormat"
		
		'span Address column header across two cells
		If chkHdrSpan.Checked Then objDataGrid.Columns(3).HeaderSpan = 2 

		objDataGrid.DataKeyFieldName = "CustomerID"		

	End Sub
	
	Public Function objDataGrid_OnClick(e)
		objDataGrid.SelectedItemKeyValue = e.Instance
	End Function

	Public Function objDataGrid_ColumnFormat(v)
		Dim rt
		rt= "<image src='images/gear.gif' width='16' height='16' align='middle'>&nbsp;" & v
		objDataGrid_ColumnFormat = rt
									
	End Function

	Public Function objDataGrid_ColumnSort(e) 
		Dim newSortCol
		
		newSortCol = CInt(e.Instance)
		
		'check if last sortcolumn and switch order
		If newSortCol = SortColumn Then 
			SortOrder = Not SortOrder
		Else
			SortOrder = 0
		End If
		
		'get new sort column
		SortColumn = newSortCol
		If SortColumn=1 Then SortField = "CompanyName"
		If SortColumn=2 Then SortField = "Contact"
		If SortColumn=3 Then SortField = "Address"
	End Function

	Public Function objDataGrid_RowDataBound(RowContext) 
		If chkRowCaption.Checked And objDataGrid.CurrentRow Mod 3=0 Then
			Section = Section +1
			RowContext.CaptionStyle = "background-color:#FCE2F8;padding:5px;"
			RowContext.Caption ="<b> "& Section &". "& _
					RowContext.DataSource("CompanyName") & _
				"</b>"
			' NOTE: When using captions and RowSpan make sure that the
			' is displayed before or after the spanned cell.
			' e.g. Above we are using RowSpan = 3. To make sure our caption appears before the spanned cell we do a check on objDataGrid.CurrentRow Mod 3=0
		End If		
	End Function	
	
%>