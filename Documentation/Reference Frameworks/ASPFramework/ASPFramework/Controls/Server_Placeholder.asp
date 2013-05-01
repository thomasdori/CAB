<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server PlaceHolder

	Public Function New_ServerPlaceholder(name) 
		Set New_ServerPlaceholder = New ServerPlaceholder
			New_ServerPlaceholder.Control.Name = name
	End Function

	Page.RegisterLibrary "ServerPlaceHolder"
	
	Class ServerPlaceholder
		Public Control

		Private mObjs
		Private mVars

		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			

			Set mVars = Server.CreateObject("Scripting.Dictionary") 'non-object storage, ie. strings, integer, etc - These are stored in the viewstate
			Set mObjs = Server.CreateObject("Scripting.Dictionary") 'only objects are stored here
		End Sub

	   Private Sub Class_Terminate()
			Set mVars = Nothing
			Set mObjs = Nothing
	   End Sub

		Public Function ReadProperties(bag)
			Dim i
			Dim keys
			Dim items

			keys = Split(bag.Read("K"),"")
			items = Split(bag.Read("I"),"")
			For i = 0 to UBound(keys)
				mVars.Add keys(i),items(i)
			Next
		End Function

		Public Function WriteProperties(bag)
			bag.Write "K",Join(mVars.Keys(),"")
			bag.Write "I",Join(mVars.Items(),"")
		End Function	    

		Public Function HandleClientEvent(e)
			HandleClientEvent = False 'To signal callee the receipt ak.
		End Function					

		Public Function Add(key,ByRef item)
			SetValue key,item
		End Function
			
		Public Function SetValue(key,Byref item)
			If mVars.Exists(key) Then mVars.Remove key
			If mObjs.Exists(key) Then mObjs.Remove key
			If IsObject(item) Then
				mObjs.Add key,item
			Else
				mVars.Add key &"",item &""
			End If
		End Function
			   
		Public Default Function Render(key)
			Dim varStart	 

			varStart = Now

			If Control.Visible = False Then
				Exit Function
			End If

			If mObjs.Exists(key) Then
				mObjs.item(key)
			ElseIf mVars.Exists(key) Then
				Response.Write mVars.item(key)
			End If	   		

			Page.TraceRender varStart,Now,Control.ControlID
		End Function
			   	
	End Class
	
%>