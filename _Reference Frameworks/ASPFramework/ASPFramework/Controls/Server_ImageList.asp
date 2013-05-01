<%
	Dim ImageListLoaded	
	ImageListLoaded = True

	Public Function New_ServerImageList(name) 
		Set New_ServerImageList = New ServerImageList
			New_ServerImageList.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerImageList"
	
	Class ServerImageList
		Dim Control
		Dim Images		

		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me
			Set Images = New_ViewStateObject()
		End Sub
	     
		Public Function ReadProperties(bag)
			 Images.LoadViewState bag.Read("S")
		End Function

		Public Function WriteProperties(bag)
			bag.Write "S",Images.GetViewState()
		End Function

		Public Function HandleClientEvent(e)
			HandleClientEvent = True 'To signal callee the receipt ak.
		End Function					

		Public Function AddImage(name,src)
			Images.Add name,src
		End Function

		Public Function Clear()
			Images.Clear
		End Function

		Public Function Count()
			Count = Images.Count
		End Function

		Public Function GetImage(name)
			GetImage = Images.GetValue(name)
		End Function

		Public Function GetImageByIndex(inx)
			GetImageByIndex = Images.GetValueByIndex(inx)
		End Function

		Public Function GetImageTag(name,w,h,alt)
			GetImageTag = "<img src='"& GetImage(name) &"' width='"& w &"' height='"& h &"' alt='"& alt &"'>"
		End Function
		
		Private Function RenderImages()
			Dim i
			Dim C
			
			C = Images.Count			
			Response.Write "<script language=""JavaScript"">"+ vbCrlf
			For i = 0 to C-1
				Response.Write "cImg = new Image();cImg.src = '" + Images.GetValueByIndex(i) +"';" + vbCrLf
			Next
			Response.Write "</script>"+ vbCrlf
		End Function

		Public Default Function Render()			
			 Dim varStart	 

			 If Control.IsVisible = False Then
				Exit Function
			 End If

			 varStart = Now

			 Call RenderImages()

			 Page.TraceRender varStart,Now,Control.Name
		End Function
			   	

	End Class
	
%>