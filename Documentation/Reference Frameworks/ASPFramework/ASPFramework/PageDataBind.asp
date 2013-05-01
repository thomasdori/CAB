<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_DataList.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<% 	Page.Execute%>	
<Span Class="Caption">Page.DataBind sample<hr></Span>
<%Page.OpenForm%>
<table border=1>
	<tr valign=top><td><%txtCompanyName%></td></tr>
	<tr valign=top><td><%txtContactName%></td></tr>
	<HR>
	<%objDataList%>
</table>
<HR>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim txtCompanyName
	Dim txtContactName
	Dim objDataList		
	
	
	Public Function Page_Init()
		Set txtCompanyName = New_ServerTextBox("txtCompanyName")
		Set txtContactName  = New_ServerTextBox("txtContactName")
		Set objDataList = New_ServerDataList("objDataList")
		txtCompanyName.Control.DataTextField = "CompanyName"
		txtContactName.Control.DataTextField = "ContactName"
	End Function

	Public Function Page_Controls_Init()
		txtCompanyName.Caption = "First Name"
		txtContactName.Caption  = "Last Name"		
		
		objDataList.ItemTemplate.FunctionName  = "fncItemTemplate"
		objDataList.Control.Style = "width:100%;border-collapse:collapse;"
		objDataList.BorderWidth = 1
		objDataList.RepeatColumns = 3
		objDataList.AlternatingItemTemplate.Style = "background-color:#DDDDDD"
		objDataList.SelectedItemTemplate.Style = "background-color:#999999;color:white"		
	End Function
	
	Public Function Page_PreRender()
		Set objDataList.DataSource = GetRecordSet("Select TOP 15 CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address From Customers")
	End Function


	Public Function fncItemTemplate(ds)
			Response.Write " <A " & Page.GetEventScript("href", "Page", "SelectRow", ds.AbsolutePosition,Ds(0)) & ">" & Ds(0).Value & "</A>"
			Response.Write "    " & ds(1).Value
	End Function
	
	Public Function Page_SelectRow(e)
			objDataList.SelectedItemIndex = CInt(e.Instance)			
			Dim rs
			Dim Sql
			Sql =  "SELECT CustomerID,CompanyName,ContactName FROM Customers WHERE CustomerID = '" & e.ExtraMessage & "'"
			Set rs = GetRecordSet(Sql)
			Page.DataBind rs,Nothing
			Set rs = Nothing
	End Function


%>