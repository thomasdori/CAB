<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_PlaceHolder.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>PlaceHolder Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--#Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">PLACEHOLDER EXAMPLE</Span>
<%Page.OpenForm%>
<table border="0">
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><%ph.Render("id")%></td>
    <td valign="top"><%ph.Render("first_name")%></td>
    <td valign="top"><%ph.Render("last_name")%></td>
  </tr>
</table>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim ph
	
	
	Public Function Page_Init()
		Set ph = New_ServerPlaceHolder("txtName")
	End Function

	Public Function Page_Load()
		ph.SetValue "id","328351"
		ph.SetValue "first_name","John"
		ph.SetValue "last_name","Brown"
	End Function
	
%>