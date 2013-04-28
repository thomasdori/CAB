<!-- #include File="Helpers/ConfigurationHelper.asp" -->
<!-- #include File="Helpers/CsrfHelper.asp" -->

<%
  Dim header
  Set header = new HeaderPartial

  Class HeaderPartial
    Function getContent()
      getContent = "
      <!DOCTYPE html>
      <!--[if lt IE 7]>      <html class=""no-js lt-ie9 lt-ie8 lt-ie7""> <![endif]-->
      <!--[if IE 7]>         <html class=""no-js lt-ie9 lt-ie8""> <![endif]-->
      <!--[if IE 8]>         <html class=""no-js lt-ie9""> <![endif]-->
      <!--[if gt IE 8]><!--> <html class=""no-js""> <!--<![endif]-->

      <head>
        <meta charset=""utf-8"">
        <meta http-equiv=""X-UA-Compatible"" content=""IE=edge,chrome=1"">
        <title></title>
        <meta name=""description"" content="""">
        <meta name=""viewport"" content=""width=device-width"">
        <meta name=""title"" content=""TMF-Highway"" />
        <meta name=""copyright"" content=""Coredat Business Solutions - A-1010 Wien, Teinfaltstrasse 8 "" />
        <META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
        <meta name=""robots"" content=""noindex,nofollow"" />
        <meta http-equiv=""pragma"" content=""no-cache"" />
        <meta http-equiv=""expires"" content=""-1"" />
        <meta http-equiv=""pragma"" content=""no-cache"" />
        <meta http-equiv=""cache-control"" content=""no-cache"" />
        <meta http-equiv=""cache-control"" content=""post-check=0,pre-check=0"" />

        <!-- Place favicon.ico and apple-touch-icon.png in the root directory -->

        <script type=""text/javascript"" src=""/Assets/js/script.js""></script>
        <script type=""text/javascript"" src=""/Assets/js/jquery-1.9.1.min.js""></script>
        <link rel=""stylesheet"" type=""text/css"" href=""/Assets/css/style.css"">

        <!-- Anti ClickJacking Script -->
        <style id=""antiClickjack"">body{display: none;}</style>
        <script type=""text/javascript"">
          if(self===top){
            var antiClickjack = document.getElementById(""antiClickjack"");
            antiClickjack.parentNode.removeChild(antiClickjack);
          } else {
            top.location = self.location;
          }
        </script>
      </head>
      <body>
          <!--[if lt IE 7]>
              <p class=""chromeframe"">You are using an <strong>outdated</strong> browser. Please <a href=""http://browsehappy.com/"">upgrade your browser</a> or <a href=""http://www.google.com/chromeframe/?redirect=true"">activate Google Chrome Frame</a> to improve your experience.</p>
          <![endif]-->

          <input type=""hidden"" value=""<%outputHelper.write(Session(CsrfHelper.getParameter()))%>"" />
      "
    End Function
  End Class
%>