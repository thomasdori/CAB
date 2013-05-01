<!--#Include File = "Controls\WebControl.asp"       -->
<!--#Include File = "Controls\Server_Panel.asp"     -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Panel Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%Call Page.Execute()%>	
<Span Class="Caption">Panel Sample<hr></Span>
<%Call Page.OpenForm()%>
<table border="0" cellpadding="4">
  <tr><td><%panel%></td></tr>
  <tr><td><%panel2%></td></tr>
  <tr><td><%panel3%></td></tr>
  </tr>
</table>
<%Call Page.CloseForm()%>
</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim panel
	Dim panel2
	Dim panel3	
	
	Public Function Page_Init()
		Set panel = New_ServerPanel("panel",200,100)
		Set panel2 = New_ServerPanel("panel2",200,100)
		Set panel3 = New_ServerPanel("panel3",200,100)
		'panel2.Control.Enabled = False
	End Function

	Public Function Page_Controls_Init()						
		panel.Text = "The Flat Panel"
		panel.Mode = 1 

		panel2.Text = "The 3D Panel"
		panel2.Mode = 2 

		panel3.Text = "The Round Panel"
		panel3.Mode = 3 
	End Function

%>