<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DropDown Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">DropDown Sample</Span>
<%Page.OpenForm%>
	<%chkHideShow%> | <%chkAutoPostBack%> | <%chkListBox%><HR>
<table>
	<tr><td valign=top><%cboDropDown%></td>
		<td valign=top><%lblMessage%></td>
	</tr>
</table>	
	<HR>
	<%cmdAdd%> | <%cmdRemove%>
	
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lblMessage
	Dim cmdAdd
	Dim cmdRemove	
	Dim chkHideShow
	Dim chkAutoPostBack
	Dim chkListBox
	Dim cboDropDown
	
	
	Public Function Page_Init()
	
		'Page.AutoResetFocus = True		
		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdAdd = New_ServerLinkButton("cmdAdd")
		Set cmdRemove = New_ServerLinkButton("cmdRemove")						
		Set chkHideShow = New_ServerCheckBox("chkHideShow")
		Set chkAutoPostBack = New_ServerCheckBox("chkAutoPostBack")
		Set chkListBox= New_ServerCheckBox("chkListBox")
		Set cboDropDown = New_ServerDropDown("cboDropDown")
		
	End Function

	Public Function Page_Controls_Init()						
		cmdAdd.Text = "Add"
		cmdRemove.Text = "Remove"

		lblMessage.Control.Style = "border:1px solid blue;background-color:#EEEEEE;width:100%;font-size:8pt"		
		lblMessage.Text = "This is an Example"
		chkHideShow.Caption = "Hide/Show List"
		chkHideShow.AutoPostBack = True
		chkAutoPostBack.Caption = "DropDown AutoPostBack"
		chkAutoPostBack.AutoPostBack=True
		chkListBox.Caption = "Make it a list box"
		chkListBox.AutoPostBack = True
		
		cboDropDown.DataTextField = "TerritoryDescription"
		cboDropDown.DataValueField = "TerritoryID"
		Set cboDropDown.DataSource = GetRecordset("SELECT TerritoryID,RTRIM(TerritoryDescription) as TerritoryDescription  FROM Territories ORDER BY 2")		
		cboDropDown.DataBind() 'Loads the items collection (that will stay in the viewstate)...
		Set cboDropDown.DataSource = Nothing 'Clear
		cboDropDown.Caption = "Territory:"
		cboDropDown.CaptionCssClass = "InputCaption"
	End Function
	Public Function Page_PreRender()
		Dim x,mx
		Dim msg 
		Set msg = New StringBuilder
		mx = cboDropDown.Items.Count -1
		
		For x = 0 To mx
			If cboDropDown.Items.IsSelected(x) Then
				msg.Append cboDropDown.Items.GetText(x) & " Is Selected, Value:" &  cboDropDown.Items.GetValue(x) & "<BR>"
			End If
		Next
		lblMessage.Text  = msg.ToString()		
		lblMessage.Control.Visible = Trim(lblMessage.Text)<>""
	End Function


	Public Function chkHideShow_Click()
		cboDropDown.Control.Visible = Not cboDropDown.Control.Visible
	End Function

	Public Function chkAutoPostBack_Click()
		cboDropDown.AutoPostBack = chkAutoPostBack.Checked
	End Function

	Public Function cmdAdd_OnClick()
		cboDropDown.Items.Append cboDropDown.Items.Count,cboDropDown.Items.Count,False
	End Function

	Public Function cmdRemove_OnClick()
		cboDropDown.Items.Remove cboDropDown.Items.Count-1
	End Function

	Public Function chkListBox_Click()
		if chkListBox.Checked Then
			cboDropDown.Rows = 10
			cboDropDown.Multiple = True
		Else
			cboDropDown.Rows = 1
			cboDropDown.Multiple = False
		End If
	End Function	
	
%>