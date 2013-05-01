<html>
<head>
<title>File Uploader</title>
<script>
	var disk = new Image(25,25);
	disk.src = "images/disk.gif";
</script>
<style>
TD {
	font-family:verdana,arial;
	font-size:11px;
}
</style>
</head>
<body leftmargin="0" bgcolor="#ECE9D8" topmargin="0">
<form method="POST" action="../../Server_FileUploader.asp?UpLoadSave=on" ENCTYPE="multipart/form-data">
<input type="hidden" name="ctrl" value="<%=Request("id")%>">
<div align="center">
  <center>
    <table border="0">
      <tr>
        <td>
<table border="0" cellpadding="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" nowrap colspan="2"><font size="1"><img border="0" src="images/pixel.gif" width="10" height="5"></font></td>
  </tr>
  <tr>
    <td width="8%" nowrap><b>Select File</b></td>
    <td width="92%">
    <hr color="#000000" size="1" align="left" width="98%">
    </td>
  </tr>
</table>
        </td>
      </tr>
      <tr>
        <td><input type="file" name="file" size="20"> </td>
      </tr>
      <tr>
        <td align="right">
          <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
              <td colspan="2"><font size="1"><img border="0" src="images/pixel.gif" width="10" height="5"></font></td>
            </tr>
            <tr>
              <td><img name="pbar" border="0" src="images/pixel.gif" width="25" height="25"></td>
              <td align="right"> <input type="submit" value=" Upload  " name="cmdupload" onclick="document.images['pbar'].src='images/disk.gif';">
                <input value="Cancel" type="button" onclick="window.close();" style="width: 80px">
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </center>
</div>
</form>
</body>
</html>
