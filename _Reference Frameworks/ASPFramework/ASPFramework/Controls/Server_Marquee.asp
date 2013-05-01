<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Marquee
	Public Function New_ServerMarquee(name) 
		Set New_ServerMarquee = New ServerMarquee
			New_ServerMarquee.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerMarquee"
	
	 Class ServerMarquee
		Dim Control
		Dim Text
		Dim Direction 'down,left,right,up
		Dim Height
		Dim Width
		Dim NumberOfLoops
		Dim ScrollDelay  'Milliseconds
		Dim ScrollAmount 'Default 6 (0) 
		Dim Behavior  'alternate, scroll, slide
		Dim HSpace
		Dim VSpace
						
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Text    = ""	
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			Text		 = bag.Read("T")
			Direction	 = bag.Read("D")
			Height       = bag.Read("H")
			Width        = bag.Read("W")
			NumberOfLoops = bag.Read("L")
			ScrollDelay  = bag.Read("SL")
			ScrollAmount = bag.Read("SA")
			Behavior	 = bag.Read("B")
			HSpace		 = bag.Read("H")
			VSpace		 = bag.Read("V")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",Text
			bag.Write "D",Direction	
			bag.Write "H",Height
			bag.Write "W",Width
			
			bag.Write "L",NumberOfLoops
			bag.Write "SL",ScrollDelay
			bag.Write "SA",ScrollAmount
			bag.Write "B",Behavior
			bag.Write "H",HSpace
			bag.Write "C",VSpace
		End Function
	   
	   Public Function HandleClientEvent(e)
	   End Function			
	   
	   Public Function SetValueFromDataSource(value)
			Text = value
	   End Function
	
	   Public Function LoadFromFile()
	   End Function
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If
			 
			 varStart = Now

			 Response.Write  "<marquee id='" & Control.ControlID & "' "
			 Response.Write  IIf(Control.Style<>""," style='" & Control.Style  + "' ","") 
			 Response.Write  IIF(Control.CssClass<>""," class='" & Control.CssClass  + "' ","") 			 
			 
					If Direction	<> "" Then Response.Write " direction = '" & Direction & "' "
					If Height       <> "" Then Response.Write " height = '" & Height & "' "
					If Width        <> "" Then Response.Write " width = '" & Width & "' "
					If NumberOfLoops <> "" Then Response.Write " loop = '" & NumberOfLoops & "' "
					If ScrollDelay  <> "" Then Response.Write " scrolldelay = '" & ScrollDelay & "' "
					If ScrollAmount <> "" Then Response.Write " scrollamount = '" & ScrollAmount & "' "
					If Behavior		<> "" Then Response.Write " behavior = '" & Behavior & "' "
					If HSpace		<> "" Then Response.Write " hspace = '" & HSpace & "' "
					If VSpace		<> "" Then Response.Write " vspace = '" & VSpace & "' "
			 
			 Response.Write ">"
			 Response.Write  Text 'Encode?
			 Response.Write  "</marquee>"
					
			Page.TraceRender varStart,Now,Control.ControlID
		
		End Function

	End Class

%>