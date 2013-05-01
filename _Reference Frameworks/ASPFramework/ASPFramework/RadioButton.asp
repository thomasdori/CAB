<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_RadioButton.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<!--#Include File = "Controls\Common.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>CheckBox and CheckBoxList Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<% 	Page.Execute %>	
<Span Class="Caption">Radio Button Sample<HR></Span>
<span><br>AutoPostBack = True is used for this example...</span>

<%Page.OpenForm%>
	<%lblMessage%><HR>
	<%chkHideShow%> | <%chkTableLayOut%> | <%chkHorizontalDirection%> | <%chkShowGrid%><HR>
	<%cmdAdd%> | <%cmdRemove%> | <%cmdAddColumnOrRow%> | <%cmdRemoveColumnOrRow%>
	<HR>
	<%optRadioList%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lblMessage
	Dim cmdAdd
	Dim cmdRemove
	Dim cmdAddColumnOrRow
	Dim cmdRemoveColumnOrRow
	
	Dim chkHideShow
	Dim chkTableLayOut
	Dim chkHorizontalDirection
	Dim optRadioList
	Dim chkShowGrid
		
	Public Function Page_Init()
		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdAdd = New_ServerLinkButton("cmdAdd")
		Set cmdRemove = New_ServerLinkButton("cmdRemove")				
		Set cmdAddColumnOrRow = New_ServerLinkButton("cmdAddColumnOrRow")
		Set cmdRemoveColumnOrRow = New_ServerLinkButton("cmdRemoveColumnOrRow")				
		
		Set chkHideShow = New_ServerCheckBox("chkHideShow")
		Set chkHorizontalDirection  = New_ServerCheckBox("chkHorizontalDirection")
		Set chkTableLayOut = New_ServerCheckBox("chkTableLayOut")
		Set optRadioList = New_ServerRadioButtonList("optRadioList")
		Set chkShowGrid  = New_ServerCheckBox("chkShowGrid")
		optRadioList.AutoPostBack=true
	End Function
	Public Function Page_Controls_Init()						
		cmdAdd.Text = "Add"
		cmdRemove.Text = "Remove"

		cmdAddColumnOrRow.Text = "Add Column"
		cmdRemoveColumnOrRow.Text = "Remove Column"

		lblMessage.Control.Style = "border:1px solid blue;background-color:#EEEEEE;width:100%;font-size:8pt"		
		lblMessage.Text = "This is an Example"
		chkHideShow.Caption = "Hide/Show List"
		chkHideShow.AutoPostBack = True
		
		chkTableLayOut.Caption = "Table Layout"
		chkTableLayOut.Checked = True
		chkTableLayOut.AutoPostBack = True
		
		chkHorizontalDirection.Caption = "Horizontal Flow"
		chkHorizontalDirection.AutoPostBack=True
		
		chkShowGrid.Caption = "Show Grid"
		chkShowGrid.AutoPostBack = True
		
		optRadioList.DataTextField = "TerritoryDescription"
		optRadioList.DataValueField = "TerritoryID"
		Set optRadioList.DataSource = GetRecordset("SELECT TerritoryID,RTRIM(TerritoryDescription) AS TerritoryDescription  FROM Territories ORDER BY 2")		
				
		optRadioList.DataBind() 'Loads the items collection (that will stay in the viewstate)...
		Set optRadioList.DataSource = Nothing 'Clear
		optRadioList.RepeatColumns=4				
		
	End Function
	
	Public Function Page_PreRender()
		Dim msg 
		Set msg = New StringBuilder
		msg.Append "chkHideShow Is checked? " & chkHideShow.Checked & "<BR>"
		msg.Append "<B>RepeatColumns:  </B>" & optRadioList.RepeatColumns &  "<BR>"
		msg.Append "<B>RepeatLayOut:  </B>" & optRadioList.RepeatLayOut &  "<BR>"
		msg.Append "<B>RepeatDirection:  </B>" & optRadioList.RepeatDirection  &  "<BR>"
		msg.Append "<B>SelectedValue:  </B>" & optRadioList.Value  &  "<BR>"
		msg.Append "<B>SelectedText:  </B>" & optRadioList.Text  &  "<BR>"
		msg.Append "<HR>"		
		lblMessage.Text  = msg.ToString()
	End Function

	Public Function chkHideShow_Click()
		optRadioList.Control.Visible = Not optRadioList.Control.Visible
	End Function
	
	Public Function chkTableLayOut_Click()
		If  chkTableLayOut.Checked   Then
			optRadioList.RepeatLayout = 1			
		Else
			optRadioList.RepeatLayout = 2
		End If		
	End Function

	Public Function chkShowGrid_Click()
		If  chkShowGrid.Checked   Then
			optRadioList.BorderWidth = 1
		Else
			optRadioList.BorderWidth=0
		End If		
	End Function

	Public Function chkHorizontalDirection_Click()
		If  chkHorizontalDirection.Checked   Then
			optRadioList.RepeatDirection=1
		Else
			optRadioList.RepeatDirection = 2
		End If		
		
	End Function

	Public Function cmdAdd_OnClick()
		optRadioList.Items.Add optRadioList.Items.Count,optRadioList.Items.Count,False,-1
	End Function

	Public Function cmdRemove_OnClick()
		optRadioList.Items.Remove optRadioList.Items.Count-1
	End Function
	

	Public Function cmdAddColumnOrRow_OnClick()
			optRadioList.RepeatColumns = optRadioList.RepeatColumns  + 1
	End Function

	Public Function cmdRemoveColumnOrRow_OnClick()
		If optRadioList.RepeatColumns - 1 >0 Then
			optRadioList.RepeatColumns = optRadioList.RepeatColumns -1
		End If
	End Function

	
%>