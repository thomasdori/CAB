<!-- #include file="adovbs.inc" -->
<-- #include file="connection.asp" ->

<%
Set kyForeign = Server.CreateObject("ADOX.Key")
Set cat = Server.CreateObject("ADOX.Catalog")

'Connect the catalog
cat.ActiveConnection = strCnn
'For i = 0 To RS.Fields
Response.Write cat.Tables("tblPriority").Keys.Item(0).Type & "-<br>"
Response.Write cat.Tables("tblPriority").Keys.Count  & "-<br>"
Response.Write kyForeign.Name & "-<br>"
'Response.Write cat.Tables("tblRequest").Key.Coulmn(0).Name


'Set RS = Conn.OpenSchema(adSchemaTables)
'Do While Not RS.EOF
'  If RS("TABLE_TYPE") = "VIEW" Then
'    Response.Write RS("TABLE_NAME") & "<br>"
'  End If
'  RS.MoveNext
'Loop

'RS.Open "tblRequest", strCnn, , , adCmdTable
'For i = 0 To RS.Fields.Count - 1
'  Response.Write RS.Fields(i).Name & ": " & RS.Fields(i).Type & "<br>"
'Next
    
%>