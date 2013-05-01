<%



'Dim x
'Dim mx

'mx = Request.Form.Count
'For x=1 to mx'
'	Response.Write Request.Form.Key(x)  & "=" & Request.Form.Item(x) & "<br>"
'next

Dim o
set o = CreateObject("Scripting.FileSystemObject")
Call o.CreateTextFile(Server.MapPath("\ASPFramework") & "\text").Write(Request.BinaryRead(Request.TotalBytes))

')
%>