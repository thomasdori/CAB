<!-- #include file="header.asp" -->

<SCRIPT Language="JavaScript">
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

<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <tr>
    <td width="100%">

<h3>DataGrid Wizard</h3>

<form name="FormMain" method="POST">
  <div align="center">
    <center>
  <table border="0" cellpadding="2" cellspacing="0" width="100%">
    <tr>
      <td width="100%">
        <table border="0" width="100%" cellspacing="1" cellpadding="2">
          <tr>
            <td align="right">Page Title:
            </td>
            <td> <input type="text" name="txtPageTitle" size="40"></td>
          </tr>
          <tr>
            <td align="right" valign="top">SQL Clause:<br>
            </td>
            <td><textarea rows="4" name="txtDisplayLsitSQLClause" cols="60">SELECT * FROM <%= strTableSelected %>
</textarea></td>
          </tr>
          <tr>
            <td align="right">Data Value Fieldname:</td>
            <td><input type="text" name="txtDataValueField" size="40">
              Eg. EmployeeID</td>
          </tr>
          <tr>
            <td align="right">Data Text Fieldname:</td>
            <td><input type="text" name="txtDataTextField" size="40">
              Eg. Email</td>
          </tr>
          <tr>
            <td align="right">Data Format String:</td>
            <td><input type="text" name="txtDataFormatString" size="60"></td>
          </tr>
          <tr>
            <td align="right"></td>
            <td>Eg. &lt;A HREF='EmployeeEdit.asp?id={1}' Class='ActionLinkNormal'&gt;{0}&lt;/A></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="100%" bgcolor="#E0E0E0">
        <p align="center">&nbsp;<br>
        <input type="button" value="Generate Now" name="ButtonGenerateNow"> Results appear in the box below<br>
        &nbsp;</p>
      </td>
    </tr>
  </table>
    </center>
  </div>
</form>

<div align="center">
  <center>
<table border="0" cellpadding="2" cellspacing="0" width="500">
  <tr>
    <td width="50%">Code for: <b>DataGrid</b></td>
    <td width="50%"><input type="button" value="Highlight &amp; Copy" name="B3" onclick="HighlightAll('ID_codeForDataGrid');"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2"><textarea rows="10" wrap="off" name="codeForDataGrid" id="ID_codeForDataGrid" cols="20" style="width:100%"></textarea></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <p align="center"></td>
  </tr>
</table>

  </center>
</div>


<SCRIPT Language="VBScript">
Sub ButtonGenerateNow_onclick
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
	strHold = strHold & vbCrLf & "    sSQL = """ & Replace(document.all.txtDisplayLsitSQLClause.value, vbCrLf, "") & """"
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