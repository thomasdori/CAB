<html>
<head>
<LINK rel="stylesheet" type="text/css" href="help.css">
</head>
<body>
<h4>You have been sucessfully logged out...</h4>
<%
Request.Cookies("CLASP_SiteUser")  = ""
Request.Cookies("CLASP_SiteEmail") = ""
%>
</body>
</html>