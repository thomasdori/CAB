<!-- #include File="ConfigurationHelper.asp" -->
<!-- #include File="CsrfHelper.asp" -->
<!-- #include File="OutputHelper.asp" -->
<!-- #include File="CookieHelper.asp" -->
<!-- #include File="EncodingHelper.asp" -->
<!-- #include File="UrlHelper.asp" -->
<!-- #include File="../Libraries/stingerASP/Stinger.asp" -->

<%
	'Dependencies: ErrorHelper.asp,

	Dim view
	Set view = new ViewHelperClass

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
	                  "    <![endif]-->" &_
	                  "    <input type=""hidden"" value=""" & output.write(Session(csrf.getParameter())) & """/>"
	    End Function

		Public Function GetFooter
			GetFooter = error.getCustomErrors & error.getStingerErrors & error.getAspErrors & "</body></html>"
		End Function
	End Class
%>