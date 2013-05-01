<!-- #include file="header.asp" -->

<%
Session("DSN") = "Provider=SQLOLEDB.1;Password=isabella;Persist Security Info=True;User ID=sa;Initial Catalog=NorthWind;Data Source=localhost"
If Session("DSN") = "" Then
	Response.Write "A data source name (DSN) is required."
Else
	'Session("DSN") = Request.Form("DSN")
	Set Conn = CreateObject("ADODB.Connection")
	Conn.Open Session("DSN")
	Set objTableRS = Conn.OpenSchema(adSchemaTables)
%>
<table border="0" cellpadding="2" cellspacing="0" width="750">
<form name="frmSelectedTable" method="POST" action="setupTable.asp">
  <tr>
    <td width="36" valign="middle"><select size="20" name="tableSelected">
<%
	Do While Not objTableRS.EOF  
		Response.Write "<option value=" & objTableRS("Table_Name") & ">" & objTableRS("Table_Name") & "</option>"
		objTableRS.MoveNext
	Loop
%>
  </select> </td>
    <td width="702" valign="middle">&nbsp; <input type="submit" value="Get fields" name="GetTableFields"></td>
  </tr>
</form>
</table>
<%
End If
%>

<!-- #include file="footer.asp" -->
