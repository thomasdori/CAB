<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
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
<%gdEmployees%>	
<HR>
<a href="EmployeeEdit.asp">Add new</a>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...	
	Dim gdEmployees
		
	Public Function Page_Init()
		
		Set gdEmployees = New_ServerDataGrid("gdEmployees")		
		Page.AutoResetScrollPosition = True
	End Function

	Public Function Page_Controls_Init()						
	End Function
	
	Public Function Page_PreRender()						
		Call ShowRecords()
	End Function
	
Public Sub SetTemplate(obj)		 
	obj.AlternatingItemStyle = "background-color:#dddddd;font-size:8pt;border:1px solid black" 
	obj.ItemStyle = "font-size:8pt;border:1px solid black"
	obj.Control.Style = "width:100%;border-collapse:collapse;"
	obj.HeaderStyle = "font-weight:bold;color:white;background-color:#4682b4;font-size:10pt;font-family:tahoma;border:1px solid black"
	obj.BorderWidth = 1
	obj.BorderColor = "Black"
	obj.GenerateColumns() 'Just so it generates most of them :-)
	
	obj.AutoGenerateColumns = False  'Now lets not have it generate the columns again during render!
	obj.Columns(0).ColumnType = 2    'Modify the column settings we need.
	obj.Columns(0).DataValueField = "EmployeeID"
	obj.Columns(0).DataTextField = "Full Name"
	obj.Columns(0).DataFormatString = "<A HREF='EmployeeEdit.asp?id={1}' Class='ActionLinkNormal'>{0}</A>"

	obj.Columns(2).ColumnType = 3    'Modify the column settings we need.
	obj.Columns(2).HeaderText = "Actions"
	obj.Columns(2).HorizonalAlign = "center"
	obj.Columns(2).CellRenderFunctionName = "RenderActionColumn"
	
	
	obj.Pager.ProgressStyle = "font-size:8pt"
	obj.Pager.PrevNextStyle = "font-size:8pt"
	obj.Pager.PagerStyle    = "font-size:8pt"
	obj.AllowPaging = True
	
	
End Sub

Public Function RenderActionColumn(ds)
		Response.Write "<A HREF='EmployeeEdit.asp?id=" & ds("EmployeeID").Value & "' Class='ActionLink'>edit</A> | "
		Response.Write "<A " & Page.GetEventScript("HREF","Page","DeleteEmployee",ds("EmployeeID").Value,"") & " Class='ActionLink'>delete</a>"		
End Function

Public Function Page_DeleteEmployee(e)
	Dim sSQL
	
	'I never execute things like this (connection and data-base wise, this is for sample purposes only)
	sSQL = "DELETE FROM Employees WHERE EmployeeID = " & Replace(e.Instance,"'","''")
	OpenConnection
		mobjConnection.Execute sSQL
	CloseConnection
End Function

Public Function ShowRecords()
	Dim sSQL
	sSQL = "SELECT [LastName] + ' ' +  [FirstName] as [Full Name], [Title], [EmployeeID] FROM Employees "
	
	Set gdEmployees.DataSource = GetRecordset(sSQL)

	SetTemplate gdEmployees

End Function


		
	
%>