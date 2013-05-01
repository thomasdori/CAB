<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_TreeView.asp" -->
<!--#Include File = "DBWrapper.asp"        -->

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<TITLE>Server TreeView Example</TITLE>
	<LINK rel="stylesheet" type="text/css" href="Samples.css">
	<style>
		.tst0i { font-weight:normal;font-size:10px;font-family:tahoma}
	</style>
</head>
<!--Include File = "Home.asp"        -->
<Span Class="Caption">TreeView Sample<hr></Span>
Still working on it sample... it will be an load on-demand treeview sample.
<%
	Page.Execute 
	Page.OpenForm
		tvTree.Render
	Page.CloseForm
%>
</BODY>
</HTML>

<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim tvTree
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		Set tvTree	 = New_ServerTreeView("tvTree","")
'		Page.DebugEnabled = True
	End Function
	
	Function Page_Controls_Init()	
		Dim sSQL
		Dim rs
		tvTree.AddRoot "Customers","",""
		sSQL = "SELECT TOP 10 CustomerID,CompanyName From Customers Order By 2"
		Set rs = GetRecordSet(sSQL)
		While Not rs.EOF
			With tvTree.Nodes(0).AddNode(rs(1).Value,"","")
				.SetAttribute "ID",rs(0).Value
				.SetAttribute "X","R"		
			End With
			rs.MoveNext
		Wend
	End Function

	Function Page_Load()
	End Function
	
	
'WEBCONTROLS POSTBACK EVENT HANDLERS
	Public Function tvTree_OnClick(node) 
		If node.GetAttribute("P") <> "" Then
			Exit Function
		End If
		node.SetAttribute "P","1"
		node.IsOpened = True
		Select Case node.GetAttribute("X")
			Case "R": 
				LoadOrders node
			Case "O": 
				LoadOrdersDetails node				
			Case Else
				Exit Function
		End Select	
	End Function
	
	Function LoadOrders(node)
		Dim sSQL
		Dim rs
		
		sSQL = "SELECT OrderID,OrderDate FROM Orders Where CustomerID = '" & Replace(node.GetAttribute("ID"),"'","''") & "'"
		Set rs = GetRecordSet(sSQL)
		While Not rs.EOF
			With node.AddNode(rs(1).Value,"","")
				.SetAttribute "ID",rs(0).Value
				.SetAttribute "X","O"
			End With
			rs.MoveNext
		Wend
	End Function

	Function LoadOrdersDetails(node)
		Dim sSQL
		Dim rs
		sSQL = "SELECT P.ProductName,OD.UnitPrice,Quantity,OD.OrderID FROM [Order Details] OD Join Products P On (OD.ProductID = P.ProductID) WHERE OD.OrderID= " & Replace(node.GetAttribute("ID"),"'","''") & " Order By ProductName"
		Set rs = GetRecordSet(sSQL)
		While Not rs.EOF
			With node.AddNode(rs(0).Value,"","")
				'.SetAttribute "ID",rs(0).Value
			End With
			rs.MoveNext
		Wend		
	End Function
		
%>
