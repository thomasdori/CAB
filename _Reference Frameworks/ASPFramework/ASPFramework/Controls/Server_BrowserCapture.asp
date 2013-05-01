<%
	Session("BrowserAgent") = "On"
%>
<html>
<body onload="capture()">
<!-- Capture Browser/Client-side Variables -->
<form name="frmForm" method="post">
<input type="hidden" value="<%=Session("TransferPage")%>" name="transfer">
<input type="hidden" value="" name="useragent">
<input type="hidden" value="" name="version">
<input type="hidden" value="" name="appname">
<input type="hidden" value="" name="dom">
<input type="hidden" value="" name="screen">
</form>
<script>
	
	function capture(){
		var an	= navigator.appName;
		var ua	= navigator.userAgent;	
		var v 	= navigator.appVersion;
		var f = document.forms["frmForm"]
		
		// store values inside form variables
		// we could add more to the list such as java enabled, client time, etc
		f.version.value 	= parseInt(v);
		f.appname.value 	= an;
		f.useragent.value 	= ua;
		f.dom.value			= (document.createElement && document.appendChild && document.getElementsByTagName)? true : false;
		f.screen.value		= screen.width +"x"+ screen.height +"x"+ screen.colorDepth;
		f.submit();
	};
</script>
<noscript>
	<input type="hidden" value="true" name="noscript">
</noscript>
<iframe width="1" height="1" style="display:none;">
	<input type="hidden" value="true" name="noiframe">
</iframe>
</body>
</html>