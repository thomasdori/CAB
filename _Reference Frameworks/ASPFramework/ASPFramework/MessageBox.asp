<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_MessageBox.asp" -->

<HTML>
	<HEAD>
		<title>PAGE MESSAGE BOX & PROMPTS</title>
		<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<LINK rel="stylesheet" type="text/css" href="Samples.css">	
	</HEAD>
	<BODY id="bodyid">
<!--#Include File = "Home.asp"        -->		
<% Page.Execute%>
<Span Class="Caption">PAGE MESSAGE BOX & PROMPTS</Span>
<%Page.OpenForm%>
	<%cmdMsg%> | <%cmdAlert%> | <%cmdConfirm%> | <%cmdPrompt%>
	<hr>
	<span>Demonstrates the use of the Page.MessageBox (Alert,Confirm,Show, and Prompt) functions</span>
<%Page.CloseForm%>

<%  

	Dim cmdMsg
	Dim cmdAlert
	Dim cmdConfirm
	Dim cmdPrompt

	Public Function Page_Init()
		Set cmdMsg = New_ServerLinkButton("cmdMsg")
		Set cmdAlert = New_ServerLinkButton("cmdAlert")
		Set cmdConfirm = New_ServerLinkButton("cmdConfirm")
		Set cmdPrompt = New_ServerLinkButton("cmdPrompt")


	End Function

	Public Function Page_Controls_Init()		
		cmdMsg.Text	= "Show Message"
		cmdAlert.Text	= "Show Alert"
		cmdConfirm.Text	= "Show Confirm"
		cmdPrompt.Text	= "Show Prompt"
	End Function

	Public Function cmdMsg_OnClick()		
		MessageBox.Show "My Message","Hello, <p>You can add your <b>HTML</b> messages right here.</p>",250,100,300,200
	End Function

	Public Function cmdAlert_OnClick()		
		MessageBox.Alert "Hello User"
	End Function

	Public Function cmdConfirm_OnClick()		
		MessageBox.Confirm "Would uou like to see another message?"
	End Function


	Public Function cmdPrompt_OnClick()		
		MessageBox.Prompt "What is your name?","[Enter name here]"
	End Function


	Public Function Page_OnMessageConfirm(e) 
		Dim id
		
		id = e.Instance
		If id="OK" Then
			MessageBox.Alert "Hi There! This is another message from your computer."
		End If
	End Function
	
	Public Function Page_OnMessagePrompt(e) 
		Dim vl
		
		vl = e.Instance
		If vl<>"[Enter name here]" Then
			MessageBox.Alert "Your Name is: " & vl
		End If
	End Function

%>
	</BODY>
</HTML>
