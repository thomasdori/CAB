<%
	Dim url
	Set url = new UrlHelper

	Class UrlHelper
		Function getApplicationUrl
			Dim protocol, https, domainname, filename, fullpath, port

	       	if Request.ServerVariables("SERVER_PORT") = "80" then
			   port = ""
			else
			   port = ":" & Request.ServerVariables("SERVER_PORT")
			end if
			protocol = "http"

		    https = lcase(request.ServerVariables("HTTPS"))
		    if https <> "off" then prot = "https"
		    domainname = Request.ServerVariables("SERVER_NAME")
		    filename = Request.ServerVariables("SCRIPT_NAME")
		    fullpath = protocol & "://" & domainname & port & filename
		    filename = right(fullpath, InStr(StrReverse(fullpath), StrReverse("/")))

		    ' does not work for subdirectories
		    getApplicationUrl = Replace(fullpath, filename, "/")

		    ' todo: refactor
		    getApplicationUrl = protocol & "://" & domainname & port & "/cab"
		End Function
	End Class
%>