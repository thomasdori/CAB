<!--#Include File = "Controls\WebControl.asp"       -->
<!--#Include File = "Controls\Server_Panel.asp"     -->
<!--#Include File = "Controls\Server_Frame.asp"     -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Frame Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Call Page.Execute()%>	
<Span Class="Caption">Frames<hr></Span>
<%Call Page.OpenForm()%>
<table border="0" cellpadding="4">
  <tr>
    <td><%panel%></td>
    <td><%fraView%></td>
  </tr>
</table>
<%Call Page.CloseForm()%>
</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim panel
	Dim fraEdit
	Dim fraView
	
	Public Function Page_Init()
		Set panel = New_ServerPanel("panel",200,100)
		Set fraEdit = New_ServerFrame("fraEdit",200,50)
		Set fraView = New_ServerFrame("fraView",200,100)
	End Function

	Public Function Page_Controls_Init()						
		panel.Text = "The Flat Panel"
		panel.Mode = 2 
		panel.PanelTemplate = "RenderPanel"
		
		'fraEdit.AutoSize = True
		fraView.Caption = "Employee Area"
		fraView.InnerHTML ="<p><center><b>Anything Goes Here!</b></center></p>"
		
	End Function

%>

<%Function RenderPanel() %>
	<span>This is inside the panel</span>
	<%fraEdit.OpenFrame%>
	<table border=0>
		<tr valign=top><td>This is inside a Frame</td></tr>
	</table>
	<%fraEdit.CloseFrame%>
<%End Function%>