<!--#Include File = "Controls\WebControl.asp"-->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_DataGrid.asp" -->
<!--#Include File = "DBWrapper.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Ajax Test</title>
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	<meta name=ProgId content=VisualStudio.HTML>
	<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
</head>
<script language="javascript">
	function oButton_OnServerEvent(data) {
		clasp.getObject("gridContainer").innerHTML = data;
	}
</script>
<%Page.Execute%>
<body>
	<%Page.OpenForm%>
		Search by: <%txtSearch%>&nbsp;&nbsp;&nbsp;<%oButton%>&nbsp;<%oButton2%>
		<hr>
		<div id="gridContainer"><%oGrid%></div>
	<%Page.CloseForm%>
</body>
</html>
<%
' code behind 
	Dim oGrid
	Dim oButton
	Dim txtSearch
	Dim oButton2
	
	Public Function Page_Init() 
		Set oGrid     = New_ServerDataGrid("oGrid")
		Set oButton   = New_ServerButton("oButton")		 
		Set oButton2  = New_ServerButton("oButton2")		 
		Set txtSearch = New_ServerTextBox("txtSearch")
	End Function
	
	Public Function Page_Controls_Init()
		oButton.Text = "Submit"
		oButton.PostMode = 1
		oButton2.Text = "Do nothing..."		
	End Function	
	
	Public Function Page_PreRender()
		oButton_OnClientEvent ""
	End Function

	Public Function oButton_OnClientEvent(e) 
		Dim rs
		Set rs = GetRecordset("SELECT * FROM Products where productname like '%" & replace(txtSearch.Text,"'","''") & "%'")		
		Set oGrid.DataSource = rs
		oGrid.HeaderStyle = "font-weight:bold;color:white;background-color:Olive;height:23px"
		oGrid.AlternatingItemStyle = "background-color:beige;"
		oGrid.TableWidth = "50%"
		oGrid.Control.Style = "cursor:default;border-color:#666633;border-collapse:collapse;font-family:tahoma;font-size:10px"
		oGrid.BorderWidth = 1
		if e <> "" Then
			oGrid.Render
			Set oGrid = Nothing	
		End If		
	End Function
%>