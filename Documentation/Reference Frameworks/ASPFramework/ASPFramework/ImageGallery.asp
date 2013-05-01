<html>
<head>
<title>Image Gallery</title>
<script language="javascript">
function setImage(url){
	var CallBackObjectID = window.CallBackObjectID;
	var obj = window.opener.WebControl.all[CallBackObjectID];
	if (obj) {
		obj.cmdExec("InsertImage",url);
		window.close();
	}
}
</script>
<style>
	TD {
		font-family:verdana, arial;
	}
</style>
</head>
<body topmargin="0" leftmargin="0" marginleft="0" margintop="0">
<table border="0" width="100%">
  <tr>
    <td width="100%" bgcolor="#006699">
      <table border="0" width="100%" cellpadding="2">
        <tr>
          <td width="100%" bgcolor="#FFFFCC"><font color="#FFFFFF" size="4">&nbsp;</font><font size="4">The
      Image Gallery</font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#EEEEEE"><font face="Arial" size="2">Select an image
      form the list below:</font>
    </td>
  </tr>
  <tr>
    <td width="100%">
      <ul>
        <li><a href="javascript:setImage('images/book.gif');"><font size="2">Book One</font></a></li>
        <li><a href="javascript:setImage('images/book01.gif');"><font size="2">Book Two</font></a></li>
        <li><a href="javascript:setImage('help/images/CLASP_Framework.jpg');"><font size="2">CLASP Framework</font></a></li>
        <li><a href="javascript:setImage('help/images/CLASP_Events.jpg');"><font size="2">CLASP Events</font></a></li>
      </ul>
    </td>
  </tr>
</table>
</body>
</html>