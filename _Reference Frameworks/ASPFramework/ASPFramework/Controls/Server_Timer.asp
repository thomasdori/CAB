<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Timer

	'Helper function.
	Public Function New_ServerTimer(name,interval) 
		Set New_ServerTimer = New ServerTimer
			New_ServerTimer.Control.Name = name
			New_ServerTimer.Interval = interval
	End Function

	Page.RegisterLibrary "ServerTimer"

	Private AlreadyWriteTimerJS
	AlreadyWriteTimerJS = False

	Class ServerTimer
		Public Control
		Public Interval	' in milliseconds

		
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Interval = 1000
		End Sub	   
		
		Public Property Get Enabled
			Enabled = Control.Enabled
		End Property
		
		Public Property Let Enabled(Value)
			Control.Enabled = Value
		End Property
		
		Public Function ReadProperties(bag)			
			Interval = bag.Read("I")
		End Function

		Public Function WriteProperties(bag)
			bag.Write "I",Interval
		End Function

		Public Function HandleClientEvent(e)
			HandleClientEvent = ExecuteEventFunction(e.EventFnc)
		End Function			

		Private Function RenderTimer()	   		
			If NOT AlreadyWriteTimerJS Then
				%>
				<script>
					function claspTimer(obj){
						if (!obj) return null;
						doPostBack('OnTimer',obj);
					}						
				</script>
				<%				
			End If

			Response.Write	"<script>" + _
							"window.setTimeout(""claspTimer('"+Control.Name+"')""," & Interval &");"+ _
							"</script>"
		End Function
	   
		Public Default Function Render()
			Dim varStart
			
			If Control.Visible = False Or Control.Enabled = False Then
				Exit Function
			End If

			varStart = Now
			RenderTimer
			Page.TraceRender varStart,Now,Control.Name
		End Function

	End Class

%>