<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_DropDown.asp"    -->
<!--#Include File = "Controls\Server_Validation.asp"    -->
<!--#Include File = "Controls\Server_ImageList.asp"    -->
<!--#Include File = "Controls\Server_Panel.asp"   -->
<!--#Include File = "DBWrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<STYLE>
</STYLE>
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">Employee Edit<HR></Span>
<a href="javascript:clasp.form.doPostBack('SetVisibility', 'Panel')">Hide/Show Panel!</a>
<%Page.OpenForm%>
<%panel%>
<%Page.CloseForm%>
</BODY>
</HTML>
<%  '
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
	Dim mRecordID
		
	Public Function Page_Init()
		mRecordID = 0  'Init
		
		Set  txtFirstName	= New_ServerTextBoxEX("txtFirstName",20,20)
		Set  txtLastName	= New_ServerTextBoxEX("txtLastName",20,20)
		Set  txtTitle		= New_ServerTextBoxEX("txtTitle",20,10)
		Set  txtTitleOfCourtesy = New_ServerTextBoxEX("txtTitleOfCourtesy",4,4)
		
		'If you want to use DATE TEXT BOXES instead then use this two lines and comment out the other declarations for these controls.
		'Keep in mind that the Server Date Text Box must be configured appropriately. Make sure that the calendar folder is accessible, if not
		'I usualy place it in the root of the website and change the relative path calendar/   to the absolute /calendar/ in the Date Text box server control. This works like a charm!
		'Set  txtBirthDate	= New_ServerDateTextBox("txtHireDate")
		'Set  txtHireDate	= New_ServerDateTextBox("txtBirthDate")
		
		Set  txtBirthDate	= New_ServerTextBoxEX("txtHireDate",10,10)
		Set  txtHireDate	= New_ServerTextBoxEX("txtBirthDate",10,10)
		
		
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
		
	End Function

	Public Function Page_Controls_Init()		
		cmdSave.Text = "Save"
		cboRegion.Bind GetRecordSet("SELECT [RegionID], [RegionDescription] FROM [Northwind].[dbo].[Region]"),"RegionID","RegionDescription","",True
		txtNotes.Mode=3
		txtNotes.Rows = 5
		txtNotes.Cols = 40
		panel.Mode = 2
		panel.OverFlow = "auto" 'scroll?
		panel.PanelTemplate = "RenderPanel"
	End Function

	Public Function Page_Load()
	
		If Not Page.IsPostBack Then
			If Request.QueryString("id")<>"" Then
				mRecordID = Clng(Request.QueryString("id"))			
				Call LoadRecord()
			End if
		End If	
	End Function
	
	Public Function Page_LoadViewState()
		mRecordID = Page.ViewState.GetValue("ID") 'Get from viewstate	
	End Function
	Public Function Page_PreRender()
		Page.ViewState.Add "ID",mRecordID 'Add to viewstate :-)
	End Function

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
			Response.Redirect "EmployeeList.asp"
		End If
		
	End Function
	
	Public Function Validate()
		
		'You could use several approaches to validation. I normaly use this approach and modify the control output as appropriate
	
		'(object,strFriendlyName, intDataType, bolRequired, strDefaultValue)
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
		
		If rs.RecordCount > 0 Then
			txtFirstName.Text = rs("FirstName").Value
			txtLastName.Text  = rs("LastName").Value
			txtTitle.Text  = rs("Title").Value
			txtTitleOfCourtesy.Text  = rs("TitleOfCourtesy").Value
			txtBirthDate.Text  = rs("HireDate").Value
			txtHireDate.Text  = rs("HireDate").Value
			txtAddress.Text  = rs("Address").Value
			txtCity.Text  = rs("City").Value
			cboRegion.Items.SetSelectedByValue rs("Region").Value,True
			txtPostalCode.Text  = rs("PostalCode").Value
			txtCountry.Text  = rs("Country").Value
			txtHomePhone.Text  = rs("HomePhone").Value
			txtNotes.Text  = rs("Notes").Value		
		End If
			
	End Function

	Public Function Panel_SetVisibility(e) 
		Panel.Control.Visible = Not Panel.Control.Visible
	End Function
%>


<%Function RenderPanel() %>
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
		<tr><td colspan=2 align=right><a href="EmployeeList.asp"><HR>Back</a> | <%cmdSave%></td></tr>
	</table>
	<%mValidation%>
<%End Function%>
