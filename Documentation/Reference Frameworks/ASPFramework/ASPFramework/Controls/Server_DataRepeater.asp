<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server ServerDataRepeater
	Public Function New_ServerDataRepeater(name) 
		Set New_ServerDataRepeater = New ServerDataRepeater
			New_ServerDataRepeater.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerDataRepeater"
	
	 Class ServerDataRepeater
		Dim Control
		Dim DataSource
		
		Dim Header
		Dim ItemTemplate
		Dim AlternateItemTemplate
		Dim Footer
		Dim SeparatorTemplate

		Private fncHeader
		Private fncItemTemplate
		Private fncAlternateItemTemplate
		Private fncFooter
		Private fncSeparatorTemplate
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			

			Header = ""
			ItemTemplate  = ""
			AlternateItemTemplate  = ""
			Footer = ""
			SeparatorTemplate = ""
			
			Set fncHeader = Nothing
			Set fncItemTemplate  = Nothing
			Set fncAlternateItemTemplate  = Nothing
			Set fncFooter = Nothing
			Set fncSeparatorTemplate = Nothing
			
			Set DataSource = Nothing
			
	   End Sub

		Private Sub Class_Terminate()
			Set DataSource = Nothing
		End Sub
	   
	   	Public Function ReadProperties(bag)			
	   		Header = bag.Read("H")
	   		ItemTemplate = bag.Read("IT")
	   		AlternateItemTemplate = bag.Read("AI")
	   		Footer  = bag.Read("F")
	   		SeparatorTemplate = bag.Read("S")
		End Function
		
		Public Function WriteProperties(bag)
	   		bag.Write "H",Header
	   		bag.Write "IT",ItemTemplate
	   		bag.Write "AI",AlternateItemTemplate
	   		bag.Write "F",Footer
	   		bag.Write "S",SeparatorTemplate
		End Function
	   
	    Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunctionEX(e)
	    End Function			
	   
	    Public Default Function Render()	
				Dim alt
				Dim varStart	 
				
				varStart = Now

				If Control.IsVisible = False Then
					Exit Function
				End If
				
				If DataSource Is Nothing Then
					Exit Function
				End If
				
	    		If Header<>"" Then
	    			Set fncHeader = GetFunctionReference(Header)
	    		End If

	    		If ItemTemplate<>"" Then
	    			Set fncItemTemplate = GetFunctionReference(ItemTemplate)
	    		End If

	    		If AlternateItemTemplate<>"" Then
	    			Set fncAlternateItemTemplate = GetFunctionReference(AlternateItemTemplate)
	    		Else
	    			Set fncAlternateItemTemplate = fncItemTemplate
	    		End If

	    		If Footer<>"" Then
	    			Set fncFooter = GetFunctionReference(Footer)
	    		End If

	    		alt = 0
	    		If Not fncHeader Is Nothing Then fncHeader()
	    		
	    		While Not DataSource.EOF
	    			
	    			If alt = 0 Then
	    				If Not fncItemTemplate Is Nothing Then fncItemTemplate(DataSource)	
	    			Else
	    				If Not fncAlternateItemTemplate Is Nothing Then fncAlternateItemTemplate(DataSource)	
	    			End If	    			
	    			alt = 1 - alt	    			
	    			If Not fncSeparatorTemplate Is Nothing Then fncSeparatorTemplate()
	    				    			
	    			DataSource.MoveNext
	    		Wend
	    		
	    		If Not fncFooter Is Nothing Then fncFooter()
				
				Page.TraceRender varStart,Now,Control.ControlID
				
	    End Function

	End Class

%>