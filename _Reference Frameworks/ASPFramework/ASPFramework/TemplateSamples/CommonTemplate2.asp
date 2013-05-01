	<style>
		A {color:black;}
		A:HOVER {color:red;}		
	</style>
<%	
	PageController.SecuredHomeURL = "PublicPage.asp"
	PageController.PublicHomeURL  = "PublicPage.asp"
 	PageController.BodyAttributes = " marginheight=0 marginwidth=0 topmargin=0 leftmargin=0 " ' You could just do this in the PageController... for a global effect.
%>

<%Public Function Page_RenderHeader()%>	
	<table height=500px width=100%><tr valign=top>
	<tr height=20px><td bgcolor="#6699cc">
		<a href="AdminPage1.asp">Admin Page 1</a> | 
		<a href="AdminPage2.asp">Admin Page 2</a> |
		<a href="ManagerPage.asp">Manager Page</a> |  
		<a href="PublicPage.asp">Public Page</a> | 
		<a href="Login.asp">Login </a> | 
		<a <%=Page.GetEventScript("href","Page","Logout","","")%>>Logout</a>
	</td></tr>
	<tr><td bgcolor=#ffffcc valign=top>
<%End Function%>

<%Public Function Page_RenderFooter()%>
	<tr height=20px><td bgcolor="#6699cc">
		<a href="AdminPage1.asp">Admin Page 1</a> | 
		<a href="AdminPage2.asp">Admin Page 2</a> |
		<a href="ManagerPage.asp">Manager Page</a> |  
		<a href="PublicPage.asp">Public Page</a> | 
		<a href="Login.asp">Login </a> | 
		<a <%=Page.GetEventScript("href","Page","Logout","","")%>>Logout</a>
	</td></tr></table>
<%End Function%>
