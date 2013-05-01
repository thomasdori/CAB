<!-- #include file="header.asp" -->

<% strTableSelected = Request.Form("tableSelected") %>

<SCRIPT Language="JavaScript">

function ProcessNow(oTable) {
//  var rows = oTable.rows;
//  alert (rows);
  row = 2;
  elm1 = document.getElementById(oTable);
  document.all.codeForAddEdit.value = document.getElementById("ID_Row_"+row+"_FieldName").value;
//  alert (elm1.rows.length);
//  var rows = elm1.rows.length;
//  alert (elm1[1].ID_Row_FieldName.innerText);
}

function markerSelect(theFlag) {
  for (var i=0; i < document.FormMain.chkSelectField.length; i++) {
    document.FormMain.chkSelectField[i].checked = theFlag;
  }
}

function HighlightRow(IDName, bol) {
  elm1 = document.getElementById(IDName);
  if (bol) {
    elm1.bgColor = "#191970";
    elm1.style.color = "#FFFFFF";
  }
  else {
    elm1.style.cursor = "";
    elm1.bgColor = "";
    elm1.style.color = "";
  }
}

function EnableDisableRow(IDName) {
  elm1 = document.getElementById(IDName);
  if (elm1.checked) {
    elm1.bgColor = "tomato";
    elm1.style.color = "#FFFFFF";
  }
  else {
    elm1.bgColor = "";
    elm1.style.color = "";
  }
}

/*
Select and Copy form element script- By Dynamicdrive.com
For full source, Terms of service, and 100s DTHML scripts
Visit http://www.dynamicdrive.com
*/

//specify whether contents should be auto copied to clipboard (memory)
//Applies only to IE 4+
//0=no, 1=yes
var copytoclip=1

function HighlightAll(theField) {
  var tempval=eval("document.all."+theField);
  tempval.focus();
  tempval.select();
  if (document.all&&copytoclip==1){
    therange=tempval.createTextRange();
    therange.execCommand("Copy");
    window.status="Contents copied to clipboard";
    setTimeout("window.status=''",1800);
  }
}

</SCRIPT>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Table</title>
<link rel="stylesheet" type="text/css" href="../../includes/includes/styles.asp">
</head>

<body>

<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <tr>
    <td width="100%"><font size="4">Table:</font> <font size="4"><%=Request.Form("tableSelected")%>

<br>

<% 
'Set objTableRS = conn.OpenSchema(adSchemaTables)
'Set objFieldsRS = conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, objTableRS(Request.Form("tableSelected")), Empty))
strSQL = "SELECT * FROM " & strTableSelected
RS.Open strSQL, Session("DSN")
%>

<form name="FormMain" method="POST">
  <table id="ID_TableMain" border="0" cellpadding="2" cellspacing="0" width="100%">
    <tr>
      <td bgcolor="#800000" width="19">&nbsp;</td>
      <td width="42" bgcolor="#800000"><font color="#FFFFFF" size="3">Field Name</font></td>
      <td width="88" bgcolor="#800000"><font color="#FFFFFF" size="3">Show Name</font></td>
      <td width="108" bgcolor="#800000"><font color="#FFFFFF" size="3">Type</font></td>
      <td width="96" bgcolor="#800000"><font color="#FFFFFF" size="3">Create</font></td>
      <td width="51" bgcolor="#800000"><font color="#FFFFFF" size="3">Width</font></td>
      <td width="31" bgcolor="#800000"><font color="#FFFFFF" size="3">Req<br>
        Field</font></td>
<%
c = 0
For Each prop in RS.Fields
  c = c + 1
'show with bgcolor
x = x xor True
If x Then
%>
    <tr id="ID_Row_<%=prop.Name%>" onmouseover="HighlightRow(this.id, true);" onmouseout="HighlightRow(this.id, false);">
<% Else %>
    <tr id="ID_Row_<%=prop.Name%>" onmouseover="HighlightRow(this.id, true);" onmouseout="HighlightRow(this.id, false);">
<%
End If
%>
      <td width="19" valign="top"><input type="checkbox" name="chkSelectField_<%=c%>" value="ON"></td>
      <td width="42" valign="top"><span id="ID_Row_<%=c%>_FieldName"><%=prop.Name%></span></td>
      <td width="88" valign="top"><input id="ID_Row_<%=c%>_ShowName" type="text" name="T1" size="10" value=<%=prop.Name%>></td>
      <td width="108" valign="top"><span id="ID_Row_<%=c%>_Type"><%ShowType(prop.type)%></span></td>
      <td width="96" valign="top"><select size="1" name="D1" id="ID_Row_<%=c%>_Input">
          <option value="Text">Text</option>
          <option value="TextArea">TextArea</option>
          <option value="Button">Button</option>
          <option value="Checkbox">Checkbox</option>
          <option value="Dropdown">Drop-down</option>
        </select></td>
      <td width="51" valign="top"><select name="D2" id="ID_Row_Size" size="1">
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="10" selected>10</option>
          <option value="15">15</option>
          <option value="20">20</option>
        </select></td>
      <td width="31" valign="top"><input type="checkbox" name="chkRequiredField" value="ON"></td>
    </tr>
<%
Next
%>
  </table>
<br>
      </font>
  <div align="center">
    <center>
  <table border="0" cellpadding="2" cellspacing="0" width="100%">
    <tr>
      <td width="100%">
        <table border="0" width="100%" cellspacing="0" cellpadding="2">
          <tr>
            <td>Page Title:</td>
            <td> <input type="text" name="txtPageTitle" size="40"></td>
          </tr>
 <font size="4">
          <tr>
            <td></td>
            <td></td>
          </tr>
        </table>
          &nbsp;</td>
    </tr>
    <tr>
      <td width="100%"><b>Build pages:</b></td>
    </tr>
 <font size="4">
    <tr>
      <td width="100%" bgcolor="#E0E0E0">
        <table border="1" width="100%">
          <tr>
            <td width="100%" colspan="2"> <input type="checkbox" name="ChkBuildDataGrid" value="ON" checked> 
            </font><b>Paged
        listing</b></td>
          </tr>
          <tr>
            <td width="50%">SQL Clause:<br>
            </td>
            <td width="50%"><textarea rows="4" name="txtDisplayLsitSQLClause" cols="60">SELECT * FROM <%= strTableSelected %>
</textarea></td>
          </tr>
          <tr>
            <td width="50%">Data Value Fieldname:</td>
            <td width="50%"><input type="text" name="txtDataValueField" size="40"><br>
              Eg. EmployeeID</td>
          </tr>
          <tr>
            <td width="50%">Data Text Fieldname:&nbsp;</td>
            <td width="50%"><input type="text" name="txtDataTextField" size="40"><br>
              Eg. Email</td>
          </tr>
          <tr>
            <td width="50%">Data Format String:&nbsp;</td>
            <td width="50%"><input type="text" name="txtDataFormatString" size="60"><br>
              Eg. &lt;A HREF='EmployeeEdit.asp?id={1}' Class='ActionLinkNormal'>{0}&lt;/A></td>
          </tr>
        </table>
 <font size="4">
        <p>&nbsp;</p>
      </font>
        <p><input type="checkbox" name="ChkBuildDisplayRecord" value="ON">
        Display record<br>
        <input type="checkbox" name="ChkBuildAdd" value="ON">
        Add<br>
        <input type="checkbox" name="ChkBuildEdit" value="ON">
        Edit<br>
        <input type="checkbox" name="ChkBuildDelete" value="ON">
        Delete<br>
        <input type="checkbox" name="ChkBuildMainPage" value="ON">
        Main page&nbsp;&nbsp;&nbsp;&nbsp;
      </p>
      </td>
    </tr>
 <font size="4">
    <tr>
      <td width="100%" bgcolor="#E0E0E0">
        <p align="center"><input type="button" value="Generate Now" name="ButtonGenerateNow"></p>
      </td>
    </tr>
  </table>
    </center>
  </div>
</form>

<p>&nbsp;</p>
<div align="center">
  <center>
<table border="0" cellpadding="2" cellspacing="0" width="500">
  <tr>
    <td width="100%">
      <h4>Results:</h4>
    </td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>DataGrid</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForDataGrid" id="ID_codeForDataGrid" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_codeForDataGrid');"></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Display Record</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForDisplayRecord" id="ID_CodeForDisplayRecord" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForDisplayRecord');"></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b> Add</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForAdd" id="ID_CodeForAdd" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForAdd');"></p>
    </td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Add...Save</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForAddSave" id="ID_CodeForAddSave" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForAddSave');"></td>
  </tr>
  <tr>
    <td width="100%"><b>&nbsp;</b></td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Edit</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForEdit" id="ID_CodeForEdit" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForEdit');"></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Edit...Save</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForEditSave" id="ID_CodeForEditSave" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForEditSave');"></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Delete</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForDelete" id="ID_CodeForDelete" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForDelete');"></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Code for: <b>Main Page</b></td>
  </tr>
  <tr>
    <td width="100%"><textarea rows="10" wrap="off" name="codeForMainPage" id="ID_CodeForMainPage" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%">
      <p align="center"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_CodeForMainPage');"></td>
  </tr>
  <tr>
    <td width="100%"></td>
  </tr>
  <tr>
    <td width="100%"></td>
  </tr>
  <tr>
    <td width="100%"></td>
  </tr>
  <tr>
    <td width="100%"></td>
  </tr>
</table>

  </center>
</div>

<%
Sub ShowType(theProp)
  strType = ""
  Select Case theProp
    Case 3
      strType = "Integer"
    Case 7
      strType = "Date"
    Case 16
      strType = "TinyInt"
    Case 17
      strType = "Unsigned TinyInt"
    Case 18
      strType = "Unsigned SmallInt"
    Case 19
      strType = "Unsigned Int"
    Case 129
      strType = "Char"
    Case 131
      strType = "Numeric"
    Case 133
      strType = "Date"
    Case 135
      strType = "DateTime"
    Case 200
      strType = "Varchar"
    Case 201
      strType = "Long Varchar"
    Case 202
      strType = "VarWChar"
    Case 203
      strType = "Long VarWchar"
    Case Else
      strType = theProp
  End Select
  Response.Write strType
End Sub

%>


<SCRIPT Language="VBScript">
'declare variables

Sub XmarkerSelect(theFlag)
  msgbox (ID_TableMain.rows.length)
End Sub


Sub ButtonGenerateNow_onclick

  If document.all.ChkBuildDisplayRecord.checked Then
    BuildCodeDisplayRecord
  End If

  If document.all.ChkBuildDataGrid.checked Then
    BuildCodeDataGrid
  End If

  If document.all.ChkBuildAdd.checked Then
    BuildCodeAdd
    BuildCodeAddSave
  End If

  If document.all.ChkBuildEdit.checked Then
    BuildCodeEdit
  End If

  If document.all.ChkBuildDelete.checked Then
    BuildCodeDelete
  End If

  If document.all.ChkBuildMainPage.checked Then
    BuildCodeMainPage
  End If
End Sub

Sub BuildCodeDisplayRecord
  document.all.codeForDisplayRecord.value = strHold
End Sub

Sub BuildCodeDataGrid
'	strHold = strHold & vbCrLf & vbCrLf

	strHold = strHold & "<" & chr(37)

	strHold = strHold & vbCrLf & "Dim objDataGrid"
	strHold = strHold & vbCrLf & "Call PageController.RenderPage()"
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function Page_Configure()"
	strHold = strHold & vbCrLf & "    PageController.PageTitle = """ & document.all.txtPageTitle.value & """"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function Page_RenderForm()"
	strHold = strHold & vbCrLf & "    objDataGrid"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function Page_Init()"
	strHold = strHold & vbCrLf & "    Set gdEmployees = New_ServerDataGrid(""objDataGrid"")"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function Page_Controls_Init()"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & "Public Function Page_PreRender()"
	strHold = strHold & vbCrLf & "    	Page.AutoResetScrollPosition = False"
	strHold = strHold & vbCrLf & "    	Page.ShowRenderTime = False"
	strHold = strHold & vbCrLf & "    'PageController.PageCSSFiles = strSiteLocation & ""/includes/styles.asp"""
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function Page_Load()"
	strHold = strHold & vbCrLf & "    Call ShowRecords()"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Function ShowRecords()"
	strHold = strHold & vbCrLf & "Dim sSQL"
	strHold = strHold & vbCrLf & "    sSQL = """ & document.all.txtDisplayLsitSQLClause.value & """"
	strHold = strHold & vbCrLf & "    Set objDataGrid.DataSource = GetRecordset(sSQL)"
	strHold = strHold & vbCrLf & "    SetTemplate objDataGrid"
	strHold = strHold & vbCrLf & "End Function"

	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "Public Sub SetTemplate(obj)"
	strHold = strHold & vbCrLf & "    obj.AlternatingItemStyle = ""background-color:#F0F0F0;font-size:8pt;border:1px solid black"""
	strHold = strHold & vbCrLf & "    obj.ItemStyle = ""font-size:8pt;border:1px solid black"""
	strHold = strHold & vbCrLf & "    obj.Control.Style = ""width:100%;border-collapse:collapse;"""
	strHold = strHold & vbCrLf & "    obj.HeaderStyle = ""font-weight:bold;color:white;background-color:#4682b4;font-size:10pt;font-family:tahoma;border:1px solid black"""
	strHold = strHold & vbCrLf & "    obj.BorderWidth = 0"
	strHold = strHold & vbCrLf & "    obj.BorderColor = ""black"""
	strHold = strHold & vbCrLf & "    obj.GenerateColumns() 'Just so it generates most of them :-)"
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "    obj.AutoGenerateColumns = False  'Now lets not have it generate the columns again during render!"
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "    obj.Columns(0).ColumnType = 2    'Modify the column settings we need."
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "    obj.Columns(0).DataValueField = """ & document.all.txtDataValueField.value & """"
	strHold = strHold & vbCrLf & "    obj.Columns(0).DataTextField = """ & document.all.txtDataTextField.value & """"
	strHold = strHold & vbCrLf & "    obj.Columns(0).DataFormatString = """ & document.all.txtDataFormatString.value & """"
	strHold = strHold & vbCrLf & "    obj.Columns(0).ColumnWidth = 200"
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "    'obj.Columns(2).Visible = False"
	strHold = strHold & vbCrLf & "    'obj.Columns(2).ColumnType = 3    'Modify the column settings we need."
	strHold = strHold & vbCrLf & "    'obj.Columns(2).HeaderText = ""Actions"""
	strHold = strHold & vbCrLf & "    'obj.Columns(2).HorizonalAlign = ""center"""
	strHold = strHold & vbCrLf & "    'obj.Columns(2).CellRenderFunctionName = ""RenderActionColumn"""
	strHold = strHold & vbCrLf
	strHold = strHold & vbCrLf & "    obj.Pager.ProgressStyle = ""font-size:8pt"""
	strHold = strHold & vbCrLf & "    obj.Pager.PrevNextStyle = ""font-size:8pt"""
	strHold = strHold & vbCrLf & "    obj.Pager.PagerStyle    = ""font-size:8pt"""
	strHold = strHold & vbCrLf & "    obj.AllowPaging = True"
	strHold = strHold & vbCrLf & "    obj.Pager.PagerSize = 2"
	strHold = strHold & vbCrLf & "End Sub"

	strHold = strHold & vbCrLf & chr(37) & ">"

	document.all.codeForDataGrid.value = strHold
End Sub

Sub BuildCodeAdd
  document.all.codeForAdd.value = strHold
End Sub


Sub BuildCodeAddSave
  document.all.codeForAddSave.value = strHold
End Sub


Sub BuildCodeEdit
  document.all.codeForEdit.value = strHold
End Sub



Function ShowInputType(theRow)
  strRowType = document.getElementByID("ID_Row_" & theRow & "_Input").value

  If LCase(strRowType) = "text" Then
    ShowInputType = "<input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & document.getElementById("ID_Row_" & theRow & "_ShowName").value & Chr(34) & ">"
  End If

  If LCase(strRowType) = "dropdown" Then
    strHold = "<select size='1' name='" & document.getElementById("ID_Row_" & theRow & "_ShowName").value & Chr(34) & "'>"
    strHold = strHold & vbCrLf & Space(8) & "<option value='1'>1</option>"
    strHold = strHold & vbCrLf & Space(6) & "</select>"
    ShowInputType = strHold
  End If

End Function

</SCRIPT>

<!-- #include file="footer.asp" -->