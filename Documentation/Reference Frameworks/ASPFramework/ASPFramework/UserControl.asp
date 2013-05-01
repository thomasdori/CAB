<!--#Include File = "Controls\Server_TextBox.asp" -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_Label.asp" -->
<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::User Control Sample

	 Class LoginControl
		
		Dim Control
		Private txtLoginName
		Private txtPassword
		Private cmdLogin
		Public OnLogin	
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
				Control.ImplementsOnInit = True				
			Set txtLoginName = New ServerTextBox
			Set txtPassword  = New ServerTextBox
			Set cmdLogin     = New_ServerLinkButton("")
			OnLogin  = ""
	   End Sub

		Public Sub OnInit()
			txtLoginName.Control.Name = Control.Name  & "_txtLogin"
			txtPassword.Control.Name = Control.Name  & "_txtPassword"
			cmdLogin.Control.Name = Control.Name  & "_Login"			
			cmdLogin.Text = "Login"
			txtPassword.Mode  = 2
			Set cmdLogin.Control.Parent = Control
		End Sub
		
	   	Public Function ReadProperties(bag)		
	   		OnLogin = bag.Read("E")		   	
		End Function
		
		Public Function WriteProperties(bag)			
			bag.Write "E",OnLogin
		End Function
	   
	   Public Function HandleClientEvent(e)
			If e.TargetObject Is cmdLogin.Control Then
				If OnLogin <> "" Then					
					Set e.TargetObject = Control 'ReRoute...
					e.EventName    = OnLogin
					ExecuteEventFunctionParams OnLogin,e
				Else
					ExecuteEventFunction Control.Name & "_OnLogin" 'Pass Trough
				End If
				HandleClientEvent = True
			Else
				HandleClientEvent = False
			End If								 	
	   End Function					
	   	
	   Public  Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If
			 varStart = Now
			 txtPassword.Text = "" 'Clear always			
			%>
				<Table border=0 bordercolor='black' cellspacing=0>
					<TR><TD>Login</TD><TD><%txtLoginName%></TD></TR>
					<TR><TD>Password</TD><TD><%txtPassword%></TD></TR>
					<TR><TD colspan=2 align=right><%cmdLogin%></TD></TR>
				</Table>
			<%

		 	 Page.TraceRender varStart,Now,Control.Name		 	 
		End Function

	End Class

%>