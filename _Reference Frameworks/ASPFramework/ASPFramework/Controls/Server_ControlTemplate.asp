<%
	Public Function New_ControlClassName(name) 
		Set New_ControlClassName = New ControlTemplate
			New_ControlClassName.Control.Name = name
	End Function
	
	Page.RegisterLibrary "<your lib name>"
	
	Class ServerControlTemplate
		Dim Control
		Dim YourVariable		
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			'Set Controls = new ControlsCollection 'Only if control is a container...
			Set Control.Owner = Me 			'Only if control is a container... if not don't expose it! (they can get to it anyway)
			
			'Initialize your internal variables
			YourVariable = ""
			Control.ImplementsOnInit = True '?  If you want the OnInit to be called
			Control.ImplementsOnLoad = True '?  If you want the Onload to be called
			Control.ImplementsProcessPostBack = True 'If you want to receive a ProcessPostBack notification on postbacks for form inputs
	   End Sub
	   
	   	Public Sub OnInit()			
	   		'At this stage of the game, the viewstate is not yet loaded but all the controls
	   		'are initialized. This can be used to add dynamic controls to this control (for more complex controls)
		End Sub

	   	Public Sub OnLoad()			
			'View state is ready and controls are ready. In here you can make decisions or update calulated fields.
			'After this, the page load will be called
			
			'If you need to register any client script (js code) you can do it using
			'Page.RegisterClientScript(scriptname,scriptcode)
			'I don't have one to register a script using the source=... put it can be easily done.
		End Sub
		
	   Public Function SetValueFromDataSource(value)
		'If you support Data Binding (Page.DataBind) then use this method to capture the binded value.
		'i.e. YourValue = value
		
		'The WebControl has a property called DataTextField (Control.DataTextField) that can be used to specify
		'the name of the field to be used during the bind procedure.
		'In the page the developer has to invoke Page.DataBind(DataSourceRecordSet,Nothing) (to bind at the page level)
		'Check the code in WebControl.asp. Right now is a flat databind and is not navigating the hierarchy. Just comment out
		'a line in the WebControl.asp Page.DataBind to support hierarchical databind.
				
	   End Function
	   
	   Public Function ReadProperties(bag)
			'If the control supports viewstate (you want to read persisted data in between requests) 
			'The bag variable is a property bag object. To see its methods go to the WebControl.asp
			'For boolean values you can use ReadBoolean, for Integers ReadInt and so on.
			
			YourVariable = bag.Read("YV")
	   End Function
		
	   Public Function WriteProperties(bag)
			'If the control supports viewstate (you want to persist data in between requests) 
			bag.Write "YV",YourVariable
	   End Function

	   Public Function ProcessPostBack()
			'Use Control.ViewState  if you want to gain access to the viewstate. 
			'It will be = Nothing EnableViewState = False.
	   End Function	    
	   Public Function HandleClientEvent(e)
			'e is of type  ClientEvent (see WebControl.asp for methods and properties)
			
			HandleClientEvent = True 'To signal callee the receipt ak.
			'You can just call the "
			HandleClientEvent = ExecuteEventFunctionEX(e)
			
			'Check the ExecuteEvent functions in WebControl.asp for more ways to handle events
			
			'or you could determine the event name (e.EventName) and handle it internally.
			
	   End Function					
		
	   Private Function RenderMe()
			Response.Write "I'm rendering myself! and by the way, YourValue is " & YourValue
			'To aquire a postback event code you must use the following series of functions:
			'Page.GetEventScript(ClientTrigger, ObjectName, EventName, ByVal Instance, ExtraMessage)
			'(and all its variations  ... check them in webcontrol.asp)
			'The ClientTrigger is the javascript event, ObjectName is the ID of the object target of the event, EventName is self explanatory
			'Instance is anything you want to pass back to the page (used normally when there is more than one control or in a link so can pass the id of something)
			'The ExtraMessage is anything EXTRA that you can to pass... :-P
			
	   End Function
	   
	   Public Default Function Render()			
			 Dim varStart	 
			 			 
			 'You are responsible to handle the meaning of "visibility". It could be as simple as exiting the function
			 
			 If Control.IsVisible = False Then
				Exit Function
			 End If
			 
			 varStart = Now
			 
			 'You may want to handle the rendering process in another private function, to keep this one small
			 Call RenderMe()
			 
			 'Be a nice boy and tell us how long did it take this control to render. This is used for debugging purposes...
		 	 Page.TraceRender varStart,Now,Control.Name
		 	
		 	 'You can always add to the trace
		 	 'If you want to relate the message to the object: Page.TraceCall(Control,Message)
		 	 'Or just to add a message Page.TraceCall(Method,Message)
		 	 'To highlite the message Page.TraceImportantCall(Control,Message)
		 	 'Or to the error:Page.TraceError
		
			 'The Control object (WebControl) has a tab index that you can use to specify the tab index.
			 'If you want the framework to assign it based on the render order use 
			 'Control.TabIndex = Page.GetNextTabIndex()
		 	 
		 	 'Need to keep the vertical scroll position? then Page.AutoResetScrollPosition = True
	   End Function
			   	

	End Class
	
%>