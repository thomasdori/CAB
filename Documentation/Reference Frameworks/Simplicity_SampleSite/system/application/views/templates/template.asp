<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<title>[%page_title%] &raquo; [%meta_title%]</title>
	<link href="[%site_url%]public/css/style.css" rel="stylesheet" type="text/css" />
</head>

<body>

<div class="container">
	<div class="header">
		<%h1(config_site_name)%>
	</div>

	<div class="nav">
		<%top_nav%>
	</div>

	<div class="content">
		
		<div class="breadcrumbs">
			[%breadcrumbs%]
		</div>
		
		<h2>[%page_title%]</h2>
		
		[%show_flash_messages%]
		[%main_content%]
		
		
	</div>

	<div class="footer">
	
		<%anchor("/","&copy; " & default_site_name, "")%>
		<br />
		Last Modified on: [%page_rendered%]<br />
		Page Executed: [%page_execute%]
	</div>
		
</div>

</body>
</html>