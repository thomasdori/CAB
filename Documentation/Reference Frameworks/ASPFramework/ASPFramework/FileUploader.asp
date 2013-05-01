<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_FileUploader.asp" -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_Image.asp" -->
<!--#Include File = "Controls\Server_ImageList.asp"    -->
<!--#Include File = "Controls\Server_Panel.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<TITLE>FileUploader Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Page.Execute%>	
<Span Class="Caption">FileUploader EXAMPLE</Span>
<%Page.OpenForm%>
<table border="0">
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><%pnlPanel%></td>
    <td valign="top"><%imgViewer%></td>
  </tr>
</table>
<%Page.CloseForm%>

<%Public Function RenderPanel()%>
	  <table border="0" width="200" bgcolor="#ECE9D8" cellpadding="2">
		<tr>
		  <td width="100%" style="border-bottom: 1px solid #C0C0C0"><b>Filters:</b><font color="#800000"> <%=cmdFile.FileFilter%></font>
		  </td>
		</tr>
		<tr>
		  <td width="100%" align="right">
			<font size="2"><%cmdFile%></font>
			</td>
		</tr>
		<tr>
		  <td width="100%">
            <b>
			<font size="2">Save Attached File As:</font></b></td>
		</tr>
		<tr>
		  <td width="100%">
			<%txtName%>&nbsp;<%cmdSave%>
			</td>
		</tr>
	  </table>
<%End Function%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim cmdFile
	Dim cmdSave
	Dim txtName
	Dim imgViewer
	Dim pnlPanel
	
	
	Public Function Page_Init()
		Set txtName				= New_ServerTextBox("txtName")
		Set imgViewer			= New_ServerImage("imgViewer")
		Set pnlPanel			= New_ServerPanel("pnlPanel",100,100)		
		Set cmdSave				= New_ServerButton("cmdSave")
		Set cmdFile 			= New_ServerFileUploader("cmdFile","")
		cmdFile.TempUploadPath	= Server.MapPath("./Uploads")		
	End Function
	

	Public Function Page_Controls_Init()
		txtName.Text = "myfile.jpg"
		cmdSave.Text = "Save"
		cmdFile.Text = "Upload File"		
		pnlPanel.PanelTemplate = "RenderPanel"
		pnlPanel.Mode = 2
	
		cmdFile.FileFilter = "bmp,jpeg,jpg,gif" 	'optional, set to "" to upload any file
		cmdFile.MaxUploadSize = 1000000				'optional, set to 0 to uplaod any size
	End Function

	Public Function Page_PreRender()
		If imgViewer.ImageSrc ="" Then
			imgViewer.Control.Visible = False
		Else
			imgViewer.Control.Visible = True
		End If	
	End Function
	
	Public Function cmdFile_OnUpload()
		imgViewer.ImageSrc = "./Uploads/" + cmdFile.TempFileName
		txtName.Text = cmdFile.FileName
	End Function
	
	Public Function cmdSave_OnClick()
		If txtName.Text <> "" Then
			cmdFile.FileName = txtName.Text
			cmdFile.SaveFile(Server.MapPath("./Uploads"))
			imgViewer.ImageSrc = "./Uploads/"+txtName.Text
		End If
	End Function

%>