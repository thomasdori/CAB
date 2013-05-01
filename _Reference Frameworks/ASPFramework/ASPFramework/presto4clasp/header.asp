<!-- #include file="adovbs.inc" -->

<%
strSiteLocation = "/aspframework/presto4clasp/"
%>

<head>

<style>
body {
	color: #000;
	background-color: #FFF;
	margin-left: 0px;
	margin-right: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}

.TableSiteTop {
	width: 100%;
	height: 60px;
	border: 0px;
	vertical-align: top;
	background-color: #C0CFE2;
}

.ContentMain {
	vertical-align: top;
	width: 100%;
	border: 0px;
	vertical-align: top;
	background-color: #FFF;
	padding-left:40px;
	padding-right:40px;
}

table {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9pt;
}
</style>

<%
Dim Conn, RS
Dim bolDSNFound
Dim bolCheckForDSN

bolDSNFound = False

Set Conn = CreateObject("ADODB.Connection")
Set RS = CreateObject("ADODB.RecordSet")

If bolCheckForDSN Then
	If Session("DSN") <> "" Then
		Conn.ConnectionString = Session("DSN")
		Conn.Open
		bolDSNFound = True
	End If

	If Not bolDSNFound Then
		If Request.Form("DSN") <> "" Then
			Conn.ConnectionString = Request.Form("DSN")
			Conn.Open
		End If
	End If

	If Not Session("bolDSNFound") Then
		Response.Redirect "noDSNFound.asp"
	End If
End If

%>
</head>

<table class="TableSiteTop">
  <tr>
    <td>
<font size="5">Presto4CLASP</font>
    </td>
  </tr>
</table>

<table class="ContentMain">
	<tr>
		<td>
<p align="center"><a href="<%=strSiteLocation%>/">
Home</a>
<p>&nbsp;</p>

