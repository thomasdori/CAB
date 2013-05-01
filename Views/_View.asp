<!-- #include File="../config.asp" -->
<!-- #include File="../Helpers/EncodingHelper.asp" -->
<!-- #include File="../Helpers/CsrfHelper.asp" -->
<!-- #include File="../Helpers/CookieHelper.asp" -->
<!-- #include File="../Helpers/ErrorHelper.asp" -->
<!-- #include File="../Helpers/OutputHelper.asp" -->
<!-- #include File="../Helpers/UrlHelper.asp" -->

<%

	Dim view : Set view = new ViewHelperClass

	Class ViewHelperClass
		Public Function GetHeader()
	      GetHeader = "<!DOCTYPE html>" &_
	                  "<!--[if lt IE 7]>      <html class=""no-js lt-ie9 lt-ie8 lt-ie7""> <![endif]-->" &_
	                  "<!--[if IE 7]>         <html class=""no-js lt-ie9 lt-ie8""> <![endif]-->" &_
	                  "<!--[if IE 8]>         <html class=""no-js lt-ie9""> <![endif]-->" &_
	                  "<!--[if gt IE 8]><!--> <html class=""no-js""> <!--<![endif]-->" &_
	                  "<head>" &_
	                  "  <meta charset=""utf-8"">" &_
	                  "  <meta http-equiv=""X-UA-Compatible"" content=""IE=edge,chrome=1"">" &_
	                  "  <title></title>" &_
	                  "  <meta name=""description"" content="""">" &_
	                  "  <meta name=""viewport"" content=""width=device-width"">" &_
	                  "  <meta name=""title"" content=""TMF-Highway"" />" &_
	                  "  <meta name=""copyright"" content=""Coredat Business Solutions - A-1010 Wien, Teinfaltstrasse 8 "" />" &_
	                  "  <META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &_
	                  "  <meta name=""robots"" content=""noindex,nofollow"" />" &_
	                  "  <meta http-equiv=""pragma"" content=""no-cache"" />" &_
	                  "  <meta http-equiv=""expires"" content=""-1"" />" &_
	                  "  <meta http-equiv=""pragma"" content=""no-cache"" />" &_
	                  "  <meta http-equiv=""cache-control"" content=""no-cache"" />" &_
	                  "  <meta http-equiv=""cache-control"" content=""post-check=0,pre-check=0"" />" &_
	                  "  <!-- Place favicon.ico and apple-touch-icon.png in the root directory -->" &_
	                  "  <script type=""text/javascript"" src=""" & url.getApplicationUrl() & "/Assets/js/script.js""></script>" &_
	                  "  <script type=""text/javascript"" src=""" & url.getApplicationUrl() & "/Assets/js/jquery-1.9.1.min.js""></script>" &_
	                  "  <link rel=""stylesheet"" type=""text/css"" href=""" & url.getApplicationUrl() & "/Assets/css/style.css"">" &_
	                  "  <!-- Anti ClickJacking Script -->" &_
	                  "  <style id=""antiClickjack"">body{display: none;}</style>" &_
	                  "  <script type=""text/javascript"">" &_
	                  "    if(self===top){" &_
	                  "      var antiClickjack = document.getElementById(""antiClickjack"");" &_
	                  "      antiClickjack.parentNode.removeChild(antiClickjack);" &_
	                  "    } else {" &_
	                  "      top.location = self.location;" &_
	                  "    }" &_
	                  "  </script>" &_
	                  "</head>" &_
	                  "<body>" &_
	                  "    <!--[if lt IE 7]>" &_
	                  "        <p class=""chromeframe"">You are using an <strong>outdated</strong> browser. Please <a href=""http://browsehappy.com/"">upgrade your browser</a> or <a href=""http://www.google.com/chromeframe/?redirect=true"">activate Google Chrome Frame</a> to improve your experience.</p>" &_
	                  "    <![endif]-->"
	    End Function

	    Public Function GetToken()
	    	GetToken = " <input type=""hidden"" id=""token"" value=""" & csrf.GetToken() & """/>"
	    End Function

		Public Function GetFooter
			GetFooter = error.getCustomErrors & error.getAspErrors & "</body></html>"
		End Function
	End Class
%>

<%
	' Draft
	'Class View
	'	Public AuthorizedRoles
	'	Public RequiresAuthentication
	'	Public RequiresAuthorization
	'	Public PageTitle
	'	Public ScriptFile
	'	Public StyleFile
	'	Public ContentPartial
'
	'	Private Sub Class_Initialize
	'		'set scriptFile to Request.ServerVariables("SCRIPT_NAME") & ".js"
	'		'set styleFile to Request.ServerVariables("SCRIPT_NAME") & ".css"
	'		'set contentpartial to Request.ServerVariables("SCRIPT_NAME") & ".html"
	'	End Sub
'
	'	Public Function OpenPage()
	'		Call  eventHelper.Execute("Page_Configure")
	'	End Function
'
	'	Public Function ClosePage()
'
	'	End Function
'
	'	Public Function RenderPage()
	'		Call OpenPage()
	'		Call GetContent()
	'		Call ClosePage()
	'	End Function
'
	'End Class
%>