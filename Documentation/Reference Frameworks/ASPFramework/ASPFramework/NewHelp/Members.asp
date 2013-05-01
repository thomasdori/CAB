<!--#Include File = "../Controls/WebControl.asp"-->
<!--#Include File = "../Controls/Server_DataGrid.asp" -->
<!--Include File = "../Controls/Server_Graph.asp" -->
<!--#Include File = "../DBWrapper.asp" -->
<%Page.Execute%>
<html>
	<head>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<title>Members</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<LINK rel="stylesheet" type="text/css" href="help.css">
	</head>
	<body>
<%Page.OpenForm%>
	 Number of members registered by year and month: <%=grdMembers2.DataSource.RecordCount%>
	<%cboGraphs%>
	<%grdMembers2%><hr>
	 Number of members registered: <%=grdMembers.DataSource.RecordCount%>
	<%grdMembers%>
<%Page.CloseForm%>
	</body>
</html>


<%
Dim grdMembers
Dim grdMembers2
Dim cboGraphs

Public Function Page_Init()
	Set grdMembers = New_ServerDataGrid("grdMembers")
	Set grdMembers2 = New_ServerDataGrid("grdMembers2")
	'Set gphData 	= New_ServerGraph("gphData")
	grdMembers.Control.EnableViewState=false
End Function

Public Function Page_Controls_Init()
	Set grdMembers.DataSource = DbLayer.GetRecordSet("SELECT FirstName,LastName,CreatedDate From ClaspUser order by 3,1,2")
	Set grdMembers2.DataSource = DbLayer.GetRecordSet("SELECT Year(CreatedDate) as Year, Month(CreatedDate) As Month, Count(*) as NumUsers From ClaspUser Group By Year(CreatedDate), Month(CreatedDate) order by 1,2")
	
	grdMembers.TableStyle="width:100%"
	grdMembers.ItemStyle = "background-color:#DDDDDD"
	grdMembers.HeaderStyle = "background-color:black;color:white"

	grdMembers2.TableStyle="width:100%"
	grdMembers2.ItemStyle = "background-color:#DDDDDD"
	grdMembers2.HeaderStyle = "background-color:black;color:white"

	gphData.Mode 		= 1
	gphData.Width		= 450
	gphData.Title 	= "Users Registered by Month/Year"
	gphData.XLabel 	= "MONTHS"
	gphData.YLabel 	= "HITS"		
	
End Function

Public Function Page_Load()
	Dim rs
	Dim aLabel
	Dim aData
	Dim x
	gphData.XScale = 1
	gphData.XScaleTick = 1
	gphData.YScale = 50	
	Set rs = grdMembers2.DataSource
	Redim aData(rs.RecordCount-1)
	
	For x = 0 to UBound(aLabel)
		aData(x) = rs(2).Value
	Next
	gphData.AddRow(dData)
	gphData.SetLegend(Array("Month"))
End Function
%>