<script language="javascript">
	 function doLogOut() {
			clasp.form.doPostBack("LogOut","Page");
	 }
</script>
<table ID="Table1">
	<tr>
		<td width="200"><IMG src="images/clasp.gif"></td>
		<td width="550" align="right" valign="bottom" class="top_menu">
			| <a href="helppage.asp" class="top_menu">home</a>&nbsp;|&nbsp;
			<a href="#" onclick="ShowPage('../help/Credits.html');" class="top_menu">about</a>&nbsp;|&nbsp;
			<a href="#" onclick="ShowPage('/ASPFramework/help/OverView.htm')" class="top_menu">framework</a>&nbsp;|&nbsp;
			<a href="#" onclick="ShowPage('downloads.asp');" class="top_menu">downloads</a>&nbsp;|&nbsp;
			<a href="#" onclick="ShowPage('http://www.gotdotnet.com/workspaces/workspace.aspx?id=69b08b15-d456-4cf9-8b12-d4642ef0c22e');" class="top_menu">support</a>&nbsp;|&nbsp;
			<a href="#" onclick="ShowPage('http://www.gotdotnet.com/workspaces/messageboard/home.aspx?id=69b08b15-d456-4cf9-8b12-d4642ef0c22e');" class="top_menu">forum</a>&nbsp;|&nbsp;
			<%If Request.Cookies("CLASP_SiteUser")<>"" Then%>
				<a href="#" onclick="doLogOut()" class="top_menu">logout</a>&nbsp;|&nbsp;
			<%Else%>
				<a href="#"  onclick="ShowPage('register.asp');" class="top_menu">login</a>&nbsp;|&nbsp;
			<%End If%>
			<a href="#" onclick="ShowPage('help.asp')" class="top_menu">help</a>&nbsp;|&nbsp;			
		</td>
		<td width="50"></td>
	</tr>
</table>
<table id="Table2" cellSpacing="0" cellPadding="3" width="100%" background="" border="0"	style="BORDER-RIGHT: #3366cc 1px solid; BORDER-TOP: #3366cc 1px solid; BORDER-LEFT: #3366cc 1px solid; BORDER-BOTTOM: #3366cc 1px solid">
	<tr>
		<td background="images/tn_08.gif" width="200" style="FONT-WEIGHT:bold;COLOR:white;HEIGHT:19px"><%=mLeftSectionTitle%></td>
		<td background="images/tn_08.gif" width="" style="FONT-WEIGHT:bold;COLOR:white;HEIGHT:19px" align=right>
		<%If "" & Request.Cookies("CLASP_SiteUser") <> "" Then
			Response.Write "Welcome back " & Server.HTMLEncode( Request.Cookies("CLASP_SiteUser") )
		End If%>
		</td>
	</tr>
	