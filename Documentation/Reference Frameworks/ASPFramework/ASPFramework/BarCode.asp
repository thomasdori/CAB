<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_BarCode.asp"        -->
<HTML>
<HEAD>
<TITLE>BarCode Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<a href='http://localhost/php/phpnet.php'>php</a>
<Span Class="Caption">BarCode Sample<hr></Span>
<!--Include File = "Home.asp"        -->
<%
Dim oRecordSet
Dim oBarCode
Page.Execute
Function Page_Init() 
	Set oBarCode = New_ServerBarCode("oBarCode")
End Function
Function Page_Load()
	Page.OpenForm
	oBarCode.Text = "*CLASP 2.0*"
	oBarCode.Height=50
	oBarCode.Narrow = 1
	Response.Write "<table><tr><td>"
	oBarCode.Render
	Response.Write "</td></tr><td align=center style='font-family:Courier New;letter-spacing: 5px;text-align: center;'>" & oBarCode.Text & "</td></tr></table>"
	Page.CloseForm
 End Function

%>
</body>
</html>

