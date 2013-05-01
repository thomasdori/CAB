<!-- #include file="../includes/adovbs.inc" -->
<!-- #include file="../includes/connection.asp" -->

<%
Set objTableRS = CreateObject("ADODB.Recordset")
'Set objTableRS = conn.OpenSchema(adSchemaTables, Array("Tassel", Empty, Empty, "Table"))
Set objTableRS = conn.OpenSchema(adSchemaTables)

Do While Not objTableRS.EOF  
  Response.Write objTableRS("Table_Name") & "<br>"
  objTableRS.MoveNext
Loop
%>