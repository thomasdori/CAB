<%
bolCheckForDSN = False
%>

<!-- #include file="header.asp" -->

<table border="0" width="100%" cellspacing="0" cellpadding="2">
  <tr>
    <td width="100%">
      <form method="POST" action="default.asp">
        <table border="0" width="100%" cellspacing="0" cellpadding="2">
          <tr>
            <td>DSN</td>
            <td><input type="text" name="DSN" size="80"></td>
          </tr>
          <tr>
            <td></td>
            <td><input type="submit" value="Next &gt;&gt;" name="B1"></td>
          </tr>
        </table>
      </form>
      <p><b>The information you enter here is NOT stored on our systems
      whatsoever. We create a temporary Session object which expires when either
      when you exit this website, or close your browser.</b></p>
      <p>DSN tips</p>
      <p>You may copy and paste your connection string.</p>
      <p>For MySQL:<br>
      driver=MySQL;server=MYSQL_SERVER_IP_ADDRESS;uid=USERNAME;pwd=PASSWORD</p>
      <p>For MS SQL Server:<br>
      Provider=SQLOLEDB.1;uid=USERNAME;password=PASSWORD;Initial
      Catalog=DATABASE NAME;Data Source=MSSQLSERVER_IP_ADDRESS</td>
  </tr>
</table>

<!-- #include file="footer.asp" -->
