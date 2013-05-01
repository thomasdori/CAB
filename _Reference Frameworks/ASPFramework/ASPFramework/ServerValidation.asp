<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_FieldValidator.asp"    -->

<HTML>
	<HEAD>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<LINK rel="stylesheet" type="text/css" href="Samples.css">	
	</HEAD>
	<BODY id="bodyid">
<!--Include File = "Home.asp"        -->		
<% Page.Execute%>
		<Span Class="Caption">Field Validators<hr></Span>
		<%Page.OpenForm%>
		<%cmdValidateGroup1%> | <%cmdValidateGroup2%> | <%cmdAvoidValidation%>
		<hr>		
		<BR>Fields In Group 1
		<table border="1" bgcolor="gainsboro" width="100%">
			<tr valign="top"><td width=1% nowrap><%txtFirstName%></td><td><%reqFirstName%><%reqDate%></td></tr>
			<tr valign="top"><td><%txtLastName%></td><td><%reqLastName%><%reqNum%></td></tr>
		</table>
		<hr>
		<BR>Fields In Group 2
		<table border="1" bgcolor="gainsboro" width="100%">
			<tr valign="top"><td>DropDown: <%cboDropDown%></td><td><%reqDropDown%></td></tr>
			<tr valign="top"><td>CheckBox: <%chkG2%></td><td><%reqChk%></td></tr>
		</table>
		<hr>
		
<%Page.CloseForm%>
<%  
	Dim txtFirstName
	Dim txtLastName
	Dim cboDropDown	
	Dim chkG2
	
	Dim cmdValidateGroup1
	Dim cmdValidateGroup2
	Dim cmdAvoidValidation
	Dim reqFirstName,reqLastName
	Dim reqDropDown
	Dim reqChk	
	
	Dim reqDate
	Dim reqNum

	Public Function Page_Init()
		Set txtFirstName = New_ServerTextBox("txtFirstName")
		Set txtLastName  = New_ServerTextBox("txtLastName")
		Set cboDropDown  = New_ServerDropDown("cboDropDown")
		Set chkG2		 = New_ServerCheckBox("chkG2")
		
		Set cmdValidateGroup1 = New_ServerLinkButton("cmdValidateGroup1")
		Set cmdValidateGroup2 = New_ServerLinkButton("cmdValidateGroup2")
		Set cmdAvoidValidation = New_ServerLinkButton("cmdAvoidValidation")
		'G1
		Set reqFirstName = New_ServerRequiredFieldValidator("reqFirstName",txtFirstName,"AQUI","","V1")
		Set reqLastName  = New_ServerRequiredFieldValidator("reqLastName",txtLastName,"This is a Required Field!","mojo","V1")
		Set reqDate      = New_ServerDateFieldValidator("reqDate",txtFirstName,"V1")
		Set reqNum       = New_ServerNumericFieldValidator("reqNum",txtLastName,"V1")
		'G2
		Set reqDropDown  = New_ServerRequiredFieldValidator("reqDropDown",cboDropDown,"DropDown is required","","V2")
		Set reqChk		 = New_ServerRequiredFieldValidator("reqChk",chkG2,"CheckBox has to be on!","","V2")
		
	End Function

	Public Function Page_Controls_Init()
		Dim x
		txtFirstName.Caption = "First Name"
		txtLastName.Caption  = "Last Name"		
		cboDropDown.Items.Append "","",True
		For x = 1 To 10
			cboDropDown.Items.Append "Item " & x,x,False
		Next
		
		reqLastName.Control.Style = "color:red;font-size:12px;font-weight:bold"		
		
		cmdValidateGroup1.Text  = "Sumbit in Group 1"
		cmdValidateGroup2.Text  = "Sumbit in Group 2"
		cmdAvoidValidation.Text = "Belongs to Group 1 but does not causes validation..."
		
		cmdValidateGroup1.ValidationGroup  = "V1"
		cmdValidateGroup2.ValidationGroup  = "V2"
		cmdAvoidValidation.ValidationGroup = "V1"
		
		'Skip validation--
		cmdAvoidValidation.CausesValidation = False
		
	End Function

	Public Function Page_Load()
		'Page.RegisterEventListener txtFirstName,"onclick","alert(1)"
	End Function


	Public Function cmdValidateGroup1_OnClick()		
	End Function
	
	Public Function cmdValidateGroup2_OnClick()
	End Function


%>
	</BODY>
</HTML>
