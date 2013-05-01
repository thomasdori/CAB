<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server HTMLGenericControl

	Public Function New_ServerHTMLGenericControl(Name,Tag) 
		Set New_ServerHTMLGenericControl = New ServerHTMLGenericControl
			New_ServerHTMLGenericControl.Tag = Tag
			New_ServerHTMLGenericControl.Control.Name = Name
	End Function

	Page.RegisterLibrary "ServerHTMLGenericControl"

	 Class ServerHTMLGenericControl
		Dim Control
		Public InnerHTML
		Dim Tag			

		Private Sub Class_Initialize()			
			Set Control  = New WebControl	
			Set Control.Owner = Me
			Tag  = ""
			InnerHTML = ""
	   End Sub

	   	Public Function ReadProperties(bag)			
			InnerHTML = bag.Read("H")
			Tag	 = bag.Read("T")
		End Function
			
		Public Function WriteProperties(bag)		
			bag.Write "H",InnerHTML
			bag.Write "T",Tag
		End Function
	   
	   Public Function HandleClientEvent(e)		 	
	   End Function					
		   
	   Private Function RenderControl()
			Response.Write "<" & Tag 	
	   		Response.Write  " id='" & Control.Name & "' "
				If Control.CssClass<>"" Then Response.Write " class='" & Control.CssClass  + "' "
				If Control.Style<>"" Then Response.Write " style='" & Control.Style + "' "
				If Control.Attributes<>"" Then Response.Write " " & Control.Attributes & " "
			Response.Write  ">"				
			Response.Write InnerHTML
			Response.Write "</" & Tag & ">"
	   End Function

	
	   Public  Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If

			 varStart = Now
			 
			 Call RenderControl()
					
		 	 Page.TraceRender varStart,Now,Control.Name
		 	 
		End Function

	End Class

%>