<!--#Include File = "Controls\WebControl.asp"       -->
<!--#Include File = "Controls\Server_Form.asp"     -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_CheckBoxList.asp" -->


<html>
<head>
<TITLE>ServeForm</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</head>
<body>
<!--Include File = "Home.asp"        -->
<%Call Page.Execute()%>	
<%Call Page.OpenForm()%>
	<%Form%>
<%Call Page.CloseForm()%>
</body>
</html>

<%  
	Dim Form
	
	Public Function Page_Init()
		Set Form = New_ServerForm("Form")
	End Function

	Public Function Page_Controls_Init()						
		Form.AddTextBox "txtFirstName","First Name",10,10,1,""
		Form.AddTextBox "txtLastName","Last Name",10,10,1,""
	End Function

%>