<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Label
	Public Function New_ServerLabel(name) 
		Set New_ServerLabel = New ServerLabel
			New_ServerLabel.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerLabel"
	
	 Class ServerLabel
		Dim Control
		Dim Text
				
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Text    = ""	
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			Text = bag.Read("T")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",Text
		End Function
	   
	   Public Function HandleClientEvent(e)
	   End Function			
	   
	   Public Function SetValueFromDataSource(value)
			Text = value
	   End Function
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If
			 
			 varStart = Now

			 Response.Write  "<span id='" & Control.ControlID & "' "
			 If Control.Style<>"" Then Response.Write " style='" & control.Style  + "' "
			 If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
			 Response.Write ">"
			 Response.Write  Text
			 Response.Write  "</span>"
					
			Page.TraceRender varStart,Now,Control.Name
		
		End Function

	End Class

%>