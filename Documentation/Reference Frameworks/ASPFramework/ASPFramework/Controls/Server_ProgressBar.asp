<%
Response.Write "<script language='JavaScript' src='" + SCRIPT_LIBRARY_PATH + "progressbar/progressbar.js'></script>"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server ProgressBar
'::::::::::::::::::
'IT IS NOT FUNCTIONAL YET!!!!!!!!!!!!!!!!!!!!!!!
'::::::::::::::::::
	Public Function New_ServerProgressBar(name) 
		Set New_ServerProgressBar = New ServerProgressBar
			New_ServerProgressBar.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerProgressBar"
	
	 Class ServerProgressBar
		Dim Control
		Dim Text
		Dim Min
		Dim Max
		Dim Width
		Dim Height
		Dim mValue
				
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Text   = ""	
			Min    = 0
			Max    = 100
			Width  = "200"
			Height = "20"
			
			mValue = 0
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			Text = bag.Read("T")
			Min  = bag.Read("M")
			Max  = bag.Read("X")
			mValue = bag.Read("V")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",Text
			bag.Write "M",Min
			bag.Write "X",Max
			bag.Write "V",mValue
		End Function
	   
	   Public Function HandleClientEvent(e)
	   End Function			
	   
	   Public Function SetValueFromDataSource(v)
			mValue = v
	   End Function
	   
	   Public Property Get Value
			Value = mValue
	   End Property    

	   Public Property Let Value(v)
			mValue = v
	   End Property    
	   
	   Public Function UpdateBrowser(v)
			mValue = v
			If Control.Visible Then
				Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
					Response.Write "CLASP_PBar_" & Control.Name & ".update(" & mValue & ");"
				Response.Write "</SCRIPT>"
				If Response.Buffer Then 
					Response.Flush
				End If
			End If
	   End Function
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			 Dim Style1
			 Dim Style2
			 
			 If Control.Visible = False Then
				Exit Function

			 End If

			 varStart = Now
			 
			 Style1 = "background-color:red;width:" & mValue & ";height:" & Height
			 Style2 = "width:" & Width & ";height:" & Height
			 
			 Response.Write "<SPAN id='" & Control.Name & "'   style='" & Style1 & "'></SPAN>"
			 Response.Write "<SPAN id='" & Control.Name & "_2' style='" & Style2 & "'></SPAN>"
			 Response.Write "<SCRIPT LANGUAGE='JavaScript'>var CLASP_PBar_" & Control.Name &  " = new CLASP_ProgressBar(" & Control.Name & "," & Min & "," & Max & "," & mValue & ");</SCRIPT>"					
			 Page.TraceRender varStart,Now,Control.Name
		
		End Function

	End Class

%>