<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "UserControl.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">Composite User Controls<hr></Span>
<Span style='font-size:8pt'><BR>This user control ("LoginControl") is made up of server text boxes and server link button...!
<br>You can look at the LoginControl control in usercontrol.asp to see how easy is to build one.
</Span>

<%Page.OpenForm%>
<table border=2>
<%	x = 0
	For r = 1 to 5
%>
		<tr valign=top>
		<%For c = 1 To 2%>
			<td><%ucLogin(X).Render()
				x = x + 1
			%></td>
		<%Next%>
		</tr>
	<%	
	Next
	%>
</table>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim  ucLogin
	Dim x,c,r	
	
	Public Function Page_Init()
		Redim ucLogin(10)
		For x = 0 To 10
			Set ucLogin(x) = New LoginControl
			ucLogin(x).Control.Name = "Login"	& x
			ucLogin(x).OnLogin = "Login"
		Next
	End Function

	Public Function Page_Controls_Init()
	End Function
	
	Public Function Login(e)
'			Call Page.ShowMessage("DID YOU KNOW THAT...","Yout Clicked Login on UC: " & e.TargetObject.Name,0,0,400,200)

	End Function


%>