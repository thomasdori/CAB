<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_DropDown.asp"    -->
<!--#Include File = "Controls\Server_Validation.asp"    -->
<!--#Include File = "Controls\Server_ImageList.asp"    -->
<!--#Include File = "Controls\Server_Panel.asp"   -->
<!--#Include File = "DBWrapper.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DatGrid Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<STYLE>
		.THEADER { background-color:#DDDDDD; }
		.TSELECTEDITEM { background-color:#AAAAAA;color:red; }
		.TALTITEM  {background-color:#DDDDDD}
		.THEADESTYLE  { font-weight:bold;color:white;background-color:#777777; }
		.ActionLink  { font-size:8pt;color:black;text-decoration:none }
		.ActionLink:hover  { color:red;;text-decoration:underline }
</STYLE>
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<% 	Page.Execute %>
<Span Class="Caption">List of Employees<HR></Span>
<%Page.OpenForm%>
<table width="100%">
<tr>
	<td width=50% valign=top>
	<% Response.Write "<a style='font-weight:bold' " & Page.GetEventScript("href","Page","LoadRecord",0,-1) & " Class='ActionLink'> add new </a>"%>
	<%gdEmployees%>
	</td>
	<td valign=top>
		<b>Form Data</b>
		<%panel%>
	</td>
</tr>
</table>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  
'#::: DataGrid
	Dim gdEmployees
	
'#::: Employee Data Fields/Controls
	Dim txtFirstName
	Dim txtLastName
	Dim txtTitle
	Dim txtTitleOfCourtesy
	Dim txtBirthDate
	Dim txtHireDate

	Dim txtAddress,txtCity,cboRegion,txtPostalCode,txtCountry
	Dim txtHomePhone,txtNotes
	Dim Panel

	Dim mValidation

	Dim cmdSave	
	
'#::: Page Variables	
	Dim mRecordID
	
'#::: Page Events		
	Public Function Page_Init()		
		mRecordID = 0  'Init

		Set gdEmployees = New_ServerDataGrid("gdEmployees")		
		
		Set  txtFirstName	= New_ServerTextBoxEX("txtFirstName",20,20)
		Set  txtLastName	= New_ServerTextBoxEX("txtLastName",20,20)
		Set  txtTitle		= New_ServerTextBoxEX("txtTitle",20,10)
		Set  txtTitleOfCourtesy = New_ServerTextBoxEX("txtTitleOfCourtesy",4,4)
		Set  txtBirthDate	= New_ServerDateTextBox("txtHireDate")
		Set  txtHireDate	= New_ServerDateTextBox("txtBirthDate")
'		Set  txtBirthDate	= New_ServerTextBoxEX("txtHireDate",10,10)
'		Set  txtHireDate	= New_ServerTextBoxEX("txtBirthDate",10,10)
		Set  txtAddress		= New_ServerTextBoxEX("txtAddress",50,50)
		Set  txtCity		= New_ServerTextBoxEX("txtCity",25,20)
		Set  cboRegion		= New_ServerDropDown("cboRegion")
		Set  txtPostalCode	= New_ServerTextBoxEX("txtPostalCode",10,10)
		Set  txtCountry		= New_ServerTextBoxEX("txtCountry",20,20)
		Set  txtHomePhone	= New_ServerTextBoxEX("txtHomePhone",15,15)
		Set  txtNotes		= New_ServerTextBox("txtNotes")
		Set cmdSave = New_ServerLinkButton("cmdSave")
		Set mValidation = New ServerValidationSummary
		Set panel = New_ServerPanel("Panel",500,400)
		Page.AutoResetScrollPosition = True
	End Function

	Public Function Page_Controls_Init()						
		cmdSave.Text = "Save"
		cboRegion.Bind GetRecordSet("SELECT [RegionID], [RegionDescription] FROM [Northwind].[dbo].[Region]"),"RegionID","RegionDescription","",True
		txtNotes.Mode=3
		txtNotes.Rows = 5
		txtNotes.Cols = 40
		panel.Mode = 2
		'panel.OverFlow = "auto" 'scroll?
		panel.PanelTemplate = "RenderPanel"
	End Function

	Public Function Page_LoadViewState()
		mRecordID = Page.ViewState.GetValue("ID") 'Get from viewstate	
	End Function

	Public Function Page_SaveViewState()
		Page.ViewState.Add "ID",mRecordID 'To viewstate I say!
	End Function
	
	Public Function Page_PreRender()						
		Call ShowRecords()
	End Function
	
	Public Function RenderActionColumn(ds)
		Response.Write "<a " & Page.GetEventScript("href","Page","LoadRecord",ds("EmployeeID").Value,ds.AbsolutePosition) & " Class='ActionLink'>edit</a> | "
		Response.Write "<a " & Page.GetEventScript("href","Page","DeleteRecord",ds("EmployeeID").Value,ds.AbsolutePosition) & " Class='ActionLink'>delete</a>"		
	End Function

	Public Function Page_LoadRecord(e)
		mRecordID = CLng(e.Instance)
		gdEmployees.SelectedItemIndex = CLng(e.ExtraMessage)
		LoadRecord()
	End Function

	Public Function Page_DeleteRecord(e)
		Dim sSQL
		
		'I never execute things like this (connection and data-base wise, this is for sample purposes only)
		sSQL = "DELETE FROM Employees WHERE EmployeeID = " & Replace(e.Instance,"'","''")
		DBLayer.ExecuteSQL sSQL
	End Function

'#::: Page Function
	Public Function ShowRecords()
		Dim sSQL
		sSQL = "SELECT [LastName] + ' ' +  [FirstName] as [Full Name], [Title], [EmployeeID] FROM Employees "
		
		Set gdEmployees.DataSource = GetRecordset(sSQL)

		SetTemplate gdEmployees

	End Function


	Public Sub SetTemplate(obj)		 
		obj.AlternatingItemStyle = "background-color:#dddddd;font-size:8pt;border:1px solid black" 
		obj.ItemStyle = "font-size:8pt;border:1px solid black"
		obj.SelectedItemStyle = "background-color:#ffaabb;font-size:8pt;border:1px solid black" 
		obj.Control.Style = "width:100%;border-collapse:collapse;"
		obj.HeaderStyle = "font-weight:bold;color:white;background-color:#4682b4;font-size:10pt;font-family:tahoma;border:1px solid black"
		obj.BorderWidth = 1
		obj.BorderColor = "Black"
		obj.GenerateColumns() 'Just so it generates most of them :-)
		
		obj.AutoGenerateColumns = False  'Now lets not have it generate the columns again during render!
		obj.Columns(2).ColumnType = 3    'Modify the column settings we need.
		obj.Columns(2).HeaderText = "Actions"
		obj.Columns(2).HorizonalAlign = "center"
		obj.Columns(2).CellRenderFunctionName = "RenderActionColumn"
		
		
		obj.Pager.ProgressStyle = "font-size:8pt"
		obj.Pager.PrevNextStyle = "font-size:8pt"
		obj.Pager.PagerStyle    = "font-size:8pt"
		obj.AllowPaging = True
			
	End Sub

	Public Function cmdSave_OnClick()
		Dim sSQL
		Dim rs
		
		Call Validate()
		If Not mValidation.IsValid Then
			Exit Function
		End If
		
		sSQL = "SELECT * FROM Employees Where EmployeeID=" & mRecordID
		Set rs = DBLayer.GetRecordSet(sSQL)
		
		If rs.RecordCount = 0 Then
			rs.AddNew
		End If
		
		rs("FirstName").Value	= txtFirstName.Text
		rs("LastName").Value    = txtLastName.Text
		rs("Title").Value		=	txtTitle.Text
		rs("TitleOfCourtesy").Value	= txtTitleOfCourtesy.Text 
		rs("BirthDate").Value	= txtBirthDate.Text
		rs("HireDate").Value	= txtHireDate.Text
		rs("Address").Value		= txtAddress.Text
		rs("City").Value		= txtCity.Text
		rs("Region").Value		= cboRegion.Value
		rs("PostalCode").Value	= txtPostalCode.Text
		rs("Country").Value		= txtCountry.Text
		rs("HomePhone").Value	= txtHomePhone.Text
		rs("Notes").Value		= txtNotes.Text
		
		
		DBLayer.UpdateRecordSet rs

		If mValidation.IsValid Then
			Response.Write "Record Saved..."
		End If
	End Function
	
	Public Function Validate()
		mValidation.Validate 	txtFirstName,"First Name", VALIDATION_DATATYPE_STRING, True , ""
		mValidation.Validate 	txtLastName,"Last Name", VALIDATION_DATATYPE_STRING,True , ""
		mValidation.Validate 	txtHireDate,"Hire Date", VALIDATION_DATATYPE_DATE,True , ""
		mValidation.Validate 	txtBirthDate,"Birth Date", VALIDATION_DATATYPE_DATE,True , ""		
	End Function

	Public Function LoadRecord()
		Dim sSQL
		Dim rs
		sSQL = "SELECT * FROM Employees Where EmployeeID=" & mRecordID
		Set rs = GetRecordSet(sSQL)
		If rs.RecordCount = 0 Then 
		   rs.AddNew 'In this sample, used just to clear data
		End If	
		'If rs.RecordCount > 0 Then
			txtFirstName.Text = "" & rs("FirstName").Value
			txtLastName.Text  = "" & rs("LastName").Value
			txtTitle.Text  = "" & rs("Title").Value
			txtTitleOfCourtesy.Text  = "" & rs("TitleOfCourtesy").Value
			txtBirthDate.Text  = "" & rs("HireDate").Value
			txtHireDate.Text  = "" & rs("HireDate").Value
			txtAddress.Text  = "" & rs("Address").Value
			txtCity.Text  = "" & rs("City").Value
			cboRegion.Items.SetSelectedByValue "" & rs("Region").Value,True
			txtPostalCode.Text  = "" & rs("PostalCode").Value
			txtCountry.Text  = "" & rs("Country").Value
			txtHomePhone.Text  = "" & rs("HomePhone").Value
			txtNotes.Text  = "" & rs("Notes").Value		
		'End If			
	End Function
			
%>
<%Function RenderPanel() %>
	<%mValidation%>
	<table border=0 ID="Table1">
		<tr valign=top><td>Name (First, Last) </td><td><%txtFirstName%>, <%txtLastName%></td></tr>
		<tr valign=top><td>Title</td><td><%txtTitle%></td></tr>
		<tr valign=top><td>Title of Courtesy</td><td><%txtTitleOfCourtesy%></td></tr>
		<tr valign=top><td>Birth Date</td><td><%txtBirthDate%></td></tr>
		<tr valign=top><td>Hire Date</td><td><%txtHireDate%></td></tr>
		<tr valign=top><td>Address</td><td><%txtAddress%></td></tr>
		<tr valign=top><td>City, Region, ZipCode</td><td><%txtCity%>, <%cboRegion%>&nbsp;<%txtPostalCode%></td></tr>
		<tr valign=top><td>Country</td><td><%txtCountry%></td></tr>
		<tr valign=top><td>Home Phone</td><td><%txtHomePhone%></td></tr>
		<tr valign=top><td>Notes</td><td><%txtNotes%></td></tr>
		<tr><td colspan=2 align=right><%cmdSave%></td></tr>
	</table>
<%End Function%>
