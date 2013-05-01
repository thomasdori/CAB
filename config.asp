<%
  OPTION EXPLICIT

  'On Error Resume Next

  Response.CodePage = 65001
  Response.CharSet = "utf-8"
  Response.Buffer=True
  Response.Expires=-1
  Response.CacheControl="no-store"

  Response.AddHeader "Expires", "Mon, 26 Jul 1997 05:00:00 GMT"
  Response.AddHeader "Last-Modified", Now & " GMT"
  Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
  Response.AddHeader "Pragma", "no-cache"

  ' Prevent the site from being rendered in a <frame> or <iframe>
  Response.AddHeader "X-FRAME-OPTIONS", "DENY"

  Server.ScriptTimeOut = 600
%>