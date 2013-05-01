<!--#Include File = "Controls/WebControl.asp" --> 
<!--#Include File = "Controls/Server_TextBox.asp" --> 
<!--#Include File = "Controls/Server_Panel.asp" --> 
<!--#Include File = "Controls/Server_Button.asp" --> 

<% 
Page.Execute

Class NewClass 
	Dim Control 
	private Panel 
	Private txt 
	private button
	Private Sub Class_Initialize()    
		Set Control = New WebControl    
		Set Control.Owner = Me           
		Control.ImplementsOnInit = True          
		Set panel = New_ServerPanel("Panel",500,400)    
		Set txt= New ServerTextBox 
	End Sub 

	Public Sub OnInit() 
		panel.Mode = 2 
		'panel.OverFlow = "auto" 
		panel.PanelTemplate = "RenderPanel" 
		panel.Control.Name= Control.Name & "_panel"       
		txt.Control.Name = Control.Name & "_txt" 
		Set panel.Control.Parent = Control 
		Set txt.Control.Parent = Control 
		set button = new ServerButton
		button.Control.Name = Control.Name & "_cmd" 
		set button.Control.Parent = Me.Control
		button.text = "TEST"
	End Sub 


	Public Function ReadProperties(bag)       

	End Function 
       
	Public Function WriteProperties(bag)          

	End Function    
       
	Public Default Function Render() 
		Dim varStart    
		If Control.Visible = False Then 
			Exit Function 
		End If 
		%> 
			<hr> 
			<% panel %>    
			<hr> 
			<% 
		End Function 
		%>    
<% Function RenderHTMLContent() %> 
	<Table border=0 bordercolor='black' cellspacing=0 ID="Table1"> 
	<TR><TD>Txt:</TD><TD><% txt%></TD></TR></table>
	<%button%> 	
<% 
End Function
	Public Function HandleClientEvent(e)
		Response.Write "CLICK!!!"
	End Function	
End Class 

%> 

<%
	Dim test 
	Sub Page_Init()
		Set test = new NewClass
		test.Control.Name = "test"
	End Sub
%>
<%Page.OpenForm
test
Page.CloseForm%>
