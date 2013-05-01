<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_DataList.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "DBWrapper.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Data List and CheckBoxList Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">DataList Example</Span>

<%Page.OpenForm%>
	<TABLE>
	<TR><TD Class="InputCaption">Hide/Show  </TD><TD><%chkHideShow%></TD></TR>
	<TR><TD Class="InputCaption">Show Grid  </TD><TD><%chkGridLines%></TD></TR>
	<TR><TD Class="InputCaption">Repeat Direction  </TD><TD><%cbRepeatDirection%></TD></TR>
	<TR><TD Class="InputCaption">Repeat Flow  </TD><TD><%cbRepeatLayout%></TD></TR>
	<TR><TD Class="InputCaption">Repeat Columns  </TD><TD><%cbRepeatColumns%></TD></TR>
	<TR><TD Class="InputCaption">Show Edit  </TD><TD><%chkShowFancy%></TD></TR>
	</TABLE>
	<%lblMessage%><HR>
	<HR>
	<%objDataList%>
	<HR>
	
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lblMessage	
	Dim chkHideShow
	Dim chkGridLines
	Dim chkShowFancy
	Dim objDataList
	Dim objLabelMessage
	Dim cbRepeatDirection
	Dim cbRepeatLayout
	Dim cbRepeatColumns
	Dim txtEdit
	
	Public Function Page_Init()
		Set lblMessage = New_ServerLabel("lblMessage")		
		Set chkHideShow = New_ServerCheckBox("chkHideShow")
		Set objDataList = New_ServerDataList("objDataList")		
		Set objLabelMessage = New StringBuilder
		Set cbRepeatDirection = New_ServerDropDown("cbRepeatDirection")
		Set cbRepeatLayout = New_ServerDropDown("cbRepeatLayout")
		Set cbRepeatColumns = New_ServerDropDown("cbRepeatColumns")
		Set txtEdit = New_ServerTextBox("txtEdit")		
		Set chkGridLines = New_ServerCheckBox("chkGridLines")
		Set chkShowFancy = New_ServerCheckBox("chkShowFancy")
	End Function

	Public Function Page_Controls_Init()						

		objDataList.RepeatColumns = 5 'Default
		objDataList.Control.Style = "width:100%;border-collapse:collapse;"
		objDataList.ItemTemplate.Style = "color:green"
		objDataList.SelectedItemTemplate.Style = "background-color:black;color:white"
		objDataList.EditItemTemplate.Style = "background-color:#6495ed"
		objDataList.BorderWidth = 0
		objDataList.HeaderTemplate.Style = "font-size:12pt;color:white;background-color:#6495ed"
		objDataList.FooterTemplate.Style = "font-size:12pt;color:white;background-color:#6495ed"
		lblMessage.Control.Style = "border:1px solid blue;background-color:#EEEEEE;width:100%;font-size:8pt"		
		chkHideShow.AutoPostBack = True		
		chkGridLines.AutoPostBack=True
		chkShowFancy.AutoPostBack=True
		
		objLabelMessage.Append   "This is an Example"
		objDataList.HeaderTemplate.FunctionName = "fncHeader"
		objDataList.FooterTemplate.FunctionName = "fncFooter"		
		objDataList.ItemTemplate.FunctionName  = "fncItemTemplate"
		objDataList.AlternatingItemTemplate.FunctionName  = "fncAlternateItemTemplate"
		objDataList.SelectedItemTemplate.FunctionName  = "fncSelectedItemTemplate"
		objDataList.EditItemTemplate.FunctionName  = "fncEditItemTemplate"
		
		cbRepeatDirection.Items.Append "Horizontal","1",True
		cbRepeatDirection.Items.Append "Vertical","2",False
		cbRepeatDirection.AutoPostBack = True
		cbRepeatLayout.Items.Append "Table","1",True
		cbRepeatLayout.Items.Append "Flow","2",False
		cbRepeatLayout.AutoPostBack=True
		
		cbRepeatColumns.Items.Append "1","1",False
		cbRepeatColumns.Items.Append  "5","5",True
		cbRepeatColumns.Items.Append "8","8",False
		cbRepeatColumns.Items.Append "12","12",False
		cbRepeatColumns.AutoPostBack = True
		
	End Function
	
	Public Function Page_PreRender()
		Dim x,mx
		Dim msg 
		Set msg = New StringBuilder
		Set objDataList.DataSource = GetRecordSet("Select  CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address From Customers")
		'lblMessage.Text = objLabelMessage.ToString()
	End Function

	Public Function chkHideShow_Click()
		objDataList.Control.Visible = Not objDataList.Control.Visible
	End Function
	
	Public Function chkGridLines_Click()
		objDataList.BorderWidth = IIF(chkGridLines.Checked,1,0)
	End Function
	
	Public Function fncHeader()
		Response.Write "The Header"
	End Function
	
	Public Function fncFooter()
	Response.Write "The Footer"
	End Function

	Public Function fncItemTemplate(ds)
			Response.Write Ds(0).Value
			If chkShowFancy.Checked Then
				Response.Write " <A " & Page.GetEventScript("HREF", "Page", "SelectRow", ds.AbsolutePosition,"") & ">Select</A>"
			End If
	End Function

	Public Function fncAlternateItemTemplate(ds)		
		Response.Write ds(0).Value
		If chkShowFancy.Checked Then
			Response.Write " <A " & Page.GetEventScript("HREF", "Page", "EditRow", ds.AbsolutePosition,"") & ">Edit</A>"
		End If
		
		'No much sense to use server controls here because the
		'all the page processing is completed... (this is already past-render) ViewState won't be restored/saved :-(
		'However, you can still take advantage of the controls to render HTML
		'The catch is that you have to handle the viewstate restore by yourself... see example below...
		'So, visibility and the such is not handled by the framework, you are on your own...
				
		'Dim txt
		'Set txt = New_ServerTextBox("txt" & ds(0).Value)
		'txt.Caption = ds(0).Value
		'If Page.IsPostBack Then
		'	txt.Text = Request(txt.Control.Name)
		'End If
		'txt.Render()

	End Function

	Public Function fncSelectedItemTemplate(ds)
		Response.Write "<B>" & dS(0).Value & "</B>"
	End Function

	Public Function fncEditItemTemplate(ds)
		txtEdit.Caption = "Code:"
		txtEdit.Text = ds(0).Value
		txtEdit.Render		
		If chkShowFancy.Checked Then
			Response.Write "<BR><A " & Page.GetEventScript("HREF", "Page", "CancelEditRow", "","") & ">Cancel</A>"
			Response.Write "| <A " & Page.GetEventScript("HREF", "Page", "UpdateRow", ds(0),"") & ">Update</A>"
		End If
	End Function	
	
	Public Function cbRepeatDirection_ItemChange(e)
		objDataList.RepeatDirection = CInt(cbRepeatDirection.Items.GetSelectedValue)
	End Function	

	Public Function cbRepeatLayout_ItemChange(e)	
		objDataList.RepeatLayOut = CInt(cbRepeatLayout.Items.GetSelectedValue)
	End Function	

	Public Function cbRepeatColumns_ItemChange(e)	
		objDataList.RepeatColumns = CInt(cbRepeatColumns.Items.GetSelectedValue)
	End Function	
	
	Public Function Page_SelectRow(e)
		objDataList.SelectedItemIndex = CInt(e.Instance)
	End Function

	Public Function Page_EditRow(e)
		objDataList.EditItemIndex = CInt(e.Instance)
		lblMessage.Text = "Editing row " & e.Instance
	End Function

	Public Function Page_CancelEditRow(e)
		objDataList.EditItemIndex = -1
		lblMessage.Text = "Cancel Editing row " & e.Instance
	End Function
	
	Public Function Page_UpdateRow(e)
		lblMessage.Text = "Updating row " & e.Instance
	End Function


%>

