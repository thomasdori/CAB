<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_CheckBoxList.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->

<script language="javascript">
	
	function testCallBack(data) {
		clasp.getObject("cboListContainer").innerHTML  = data;
	}

	function AjaxdoPostBack() {		
		clasp.ajax.invoke("ClickButton",true,testCallBack);		
	}
	function AjaxdoPostBack2() {		
		clasp.ajax.invoke("ClickButton2",false,testCallBack,'cmdLink',true);		
	}
	function AjaxdoPostBack3() {		
		alert( clasp.page.cmdLinkButton.invoke() ) ;
	}
	
	function AjaxdoPostBack4() {											
		//var a = 
		clasp.ajax.invoke("Page_OnClientSideEvent",true,tst,'Page',false,'AjaxActions.asp');
		//alert(a)
	}
	
	function tst(data) {
		alert(data);
	}
	function tst2() {
		window.setTimeout(AjaxdoPostBack4,1)
	}
</script>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<TITLE>Server Buttons Example</TITLE>
		<LINK rel="stylesheet" type="text/css" href="Samples.css">
	</HEAD>
	<BODY id="CLASPBody">
		<%Page.Execute%>
		<Span Class="Caption">Buttons Sample<hr></Span>
		<br><%lblMessage%>
		<%Page.OpenForm%>
			<a href="javascript:AjaxdoPostBack()">Test Ajax1</a> | 
			<a href="javascript:AjaxdoPostBack2()">Test Ajax2</a> | 
			<a href="javascript:AjaxdoPostBack3()">Test Ajax3</a> | 
			<a <%=Page.GetAjaxEventScript(cmdButton,"tst",false,"href")%>>Using Page.GetAjaxEventScript</a> | 
			<a href='javascript:tst2()'>To Anothe page?</a>
			<hr>
			Text Box:<%txtTest%>&nbsp;CheckBox:<%chkBox%>&nbsp;Drop Down:<span id="cboListContainer"><%cboList%></span>
			<br>Check Box List:<%chkList%>
						
			<%cmdLinkButton%> | <%cmdButton%> | <%cmdAdvanceButton%> | <%cmdImageButton%>
		<%Page.CloseForm%>
	</BODY>
</HTML>
<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim lblMessage
	Dim cmdLinkButton
	Dim cmdAdvanceButton	
	Dim cmdButton
	Dim cmdImageButton	
	Dim txtTest	
	Dim chkList
	Dim chkBox
	Dim cboList
	
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		'Page.DebugEnabled = True
		Set lblMessage = New_ServerLabel("lblMessage")
		Set cmdLinkButton = New_ServerLinkButton("cmdLinkButton")
		Set cmdButton = New_ServerButton("cmdButton")
		Set cmdAdvanceButton = New_ServerAdvanceButton("cmdAdvanceButton")		
		Set cmdImageButton = New_ServerImageButton("cmdImageButton")	
		
		Set txtTest = New_ServerTextBox("txtTest")
		Set chkList = New_ServerCheckBoxList("chkList")
		Set chkBox  = New_ServerCheckBox("chkBox")
		Set cboList = New_ServerDropDown("cboList")
'		Page.ViewStateMode = VIEW_STATE_MODE_CLIENT
	End Function

	Function Page_Controls_Init()	
		cmdButton.Text = "Normal Button"
		cmdLinkButton.Text = "Link Button"
		cmdAdvanceButton.Text = "<b><font color='red'>A</font><font color='maroon'>dvance</font></b> <font color='navy'>Button</font>"
		cmdImageButton.Image = "images/book.gif"
		cmdImageButton.RollOverImage = "images/clear_all.gif"
		lblMessage.Control.Style = "border:1px solid blue;background-color:#AAAAAA"
		
		cboList.DataTextField = "TerritoryDescription"
		cboList.DataValueField = "TerritoryID"
	
		Set cboList.DataSource = GetRecordset("SELECT top 12 TerritoryID,RTRIM(TerritoryDescription) as TerritoryDescription  FROM Territories ORDER BY 2")		
		cboList.DataBind()
		
		chkList.DataTextField = "TerritoryDescription"
		chkList.DataValueField = "TerritoryID"
		Set chkList.DataSource = cboList.DataSource
		chkList.RepeatColumns = 5
		chkList.dataBind()
		Set chkList.DataSource = Nothing
		Set cboList.DataSource = Nothing 'Clear
		
		
		
	End Function

	Function Page_Load()
		Page.RegisterWebControlsInClientSide
	End Function
	
	Public Function Page_Terminate()
		'Response.Clear
		'Response.Write Server.HTMLEncode(Page.Viewstate.GetViewState)
		'Response.End

	End Function
'WEBCONTROLS POSTBACK EVENT HANDLERS	
	Function cmdButton_OnClick()
		lblMessage.Text = "You Clicked a normal button"		
	End Function

	Function cmdAdvanceButton_OnClick()
		lblMessage.Text = "You Clicked an Advance button"
	End Function

	Function cmdLinkButton_OnClick()
		lblMessage.Text = "You Clicked  Link Button"
	End Function
	
	Function cmdImageButton_OnClick()
		lblMessage.Text = "You Clicked on a Image Button"
	End Function

	function cmdLinkButton_OnClientEvent(e)
		response.Write "Retrieving the results as a return value..." & Now & " julio was here"
	end function
		
	function ClickButton(e)
		cboList.Items.Add Now , cboList.Items.Count ,true,0
		cboList.Render
	End function
	
	Function ClickButton2(e)
		Response.Write "Retrieving using a sync function: " & Now
	End function
	
	Function cmdButton_OnClientEvent(e)
		response.Write txtTest.text & vbNewLine & chkBox.Checked & vbNewLine & cboList.Text & "(" & cboList.value & ")"
	End Function
	
'SUPPORTING FUNCTIONS (IF ANY)
%>
