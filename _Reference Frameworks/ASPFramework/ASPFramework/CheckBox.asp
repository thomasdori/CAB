<!--#Include File = "Controls/WebControl.asp"-->
<!--#Include File = "Controls/Server_Button.asp" -->
<!--#Include File = "Controls/Server_CheckBox.asp" -->
<!--#Include File = "Controls/Server_CheckBoxList.asp" -->
<!--#Include File = "Controls/Server_Label.asp"    -->
<!--#Include File = "Controls/Common.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->
<html>
<head>
	<title>CheckBox Example</title>
	<link rel="stylesheet" type="text/css" href="Samples.css">
</head>
<body>
<!--Include File = "Home.asp"    -->
<% 
	Call Page.Execute
	Call Page.OpenForm
%>
	<%lblMessage%><HR>
	<%chkHideShow%> | <%chkTableLayOut%> | <%chkHorizontalDirection%> | <%chkShowGrid%><HR>
	<%cmdAdd%> | <%cmdRemove%> | <%cmdAddColumnOrRow%> | <%cmdRemoveColumnOrRow%>
	<HR>
	<%chkCheckBoxList%>
<%	Call Page.CloseForm%>
</body>
</html>
<%  
':: CONTROL DECLARATIONS
	Dim lblMessage
	Dim cmdAdd
	Dim cmdRemove
	Dim cmdAddColumnOrRow
	Dim cmdRemoveColumnOrRow
	
	Dim chkHideShow
	Dim chkTableLayOut
	Dim chkHorizontalDirection
	Dim chkCheckBoxList
	Dim chkShowGrid
	
':: PAGE EVENTS
	Public Function Page_Init()
   		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdAdd = New_ServerLinkButton("cmdAdd")
		Set cmdRemove = New_ServerLinkButton("cmdRemove")				
		Set cmdAddColumnOrRow = New_ServerLinkButton("cmdAddColumnOrRow")
		Set cmdRemoveColumnOrRow = New_ServerLinkButton("cmdRemoveColumnOrRow")				
		
		Set chkHideShow = New_ServerCheckBox("chkHideShow")
		Set chkHorizontalDirection  = New_ServerCheckBox("chkHorizontalDirection")
		Set chkTableLayOut = New_ServerCheckBox("chkTableLayOut")
		Set chkCheckBoxList = New_ServerCheckBoxList("chkCheckBoxList")
		Set chkShowGrid  = New_ServerCheckBox("chkShowGrid")
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
		
		'Test Caching:
		Dim rs
		Set rs = Nothing
		If Application("CHKTST") = "" Then
			Set rs = GetRecordset("SELECT TerritoryID,rtrim(TerritoryDescription) As TerritoryDescription FROM Territories ORDER BY 2 ASC")					
			Call CacheListItemsCollection(rs,"TerritoryDescription","TerritoryID","CHKTST",False,"")
			Set rs = Nothing
		End If
		SetListItemsCollectionFromCache chkCheckBoxList,"CHKTST"		
		chkCheckBoxList.RepeatColumns=4
	End Function
	
	Public Function Page_PreRender()
		Dim x,mx
		Dim msg 
		Set msg = New StringBuilder
		msg.Append "chkHideShow Is checked? " & chkHideShow.Checked & "<BR>"
		msg.Append "<B>RepeatColumns:  </B>" & chkCheckBoxList.RepeatColumns &  "<BR>"
		msg.Append "<B>RepeatLayOut:  </B>" & chkCheckBoxList.RepeatLayOut &  "<BR>"
		msg.Append "<B>RepeatDirection:  </B>" & chkCheckBoxList.RepeatDirection  &  "<BR>"
		msg.Append "<HR>"
		mx = chkCheckBoxList.Items.Count -1
				
		For x = 0 To mx
			If chkCheckBoxList.Items.IsSelected(x) Then
				msg.Append chkCheckBoxList.Items.GetText(x) & " Is Selected, Value:" &  chkCheckBoxList.Items.GetValue(x) & "<BR>"
			End If
		Next
		lblMessage.Text  = msg.ToString()
	End Function

'WEBCONTROLS POSTBACK EVENT HANDLERS	
	Public Function chkHideShow_Click()
		chkCheckBoxList.Control.Visible = Not chkCheckBoxList.Control.Visible
	End Function
	
	Public Function chkTableLayOut_Click()
		If  chkTableLayOut.Checked   Then
			chkCheckBoxList.RepeatLayout = 1			
		Else
			chkCheckBoxList.RepeatLayout = 2
		End If		
	End Function

	Public Function chkShowGrid_Click()
		If  chkShowGrid.Checked   Then
			chkCheckBoxList.BorderWidth = 1
		Else
			chkCheckBoxList.BorderWidth=0
		End If		
	End Function

	Public Function chkHorizontalDirection_Click()
		If  chkHorizontalDirection.Checked   Then
			chkCheckBoxList.RepeatDirection=1
		Else
			chkCheckBoxList.RepeatDirection = 2
		End If		
		
	End Function

	Public Function cmdAdd_OnClick()
		chkCheckBoxList.Items.Append chkCheckBoxList.Items.Count,chkCheckBoxList.Items.Count,False
	End Function

	Public Function cmdRemove_OnClick()
		chkCheckBoxList.Items.Remove chkCheckBoxList.Items.Count-1
	End Function
	

	Public Function cmdAddColumnOrRow_OnClick()
			chkCheckBoxList.RepeatColumns = chkCheckBoxList.RepeatColumns  + 1
	End Function

	Public Function cmdRemoveColumnOrRow_OnClick()
		If chkCheckBoxList.RepeatColumns - 1 >0 Then
			chkCheckBoxList.RepeatColumns = chkCheckBoxList.RepeatColumns -1
		End If
	End Function

	
%>