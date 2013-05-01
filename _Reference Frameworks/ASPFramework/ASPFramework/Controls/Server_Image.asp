<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Image

	Public Function New_ServerImage(name) 
		Set New_ServerImage = New ServerImage
			New_ServerImage.Control.Name = name
	End Function

	Page.RegisterLibrary "ServerImage"

	 Class ServerImage
		Dim Control
		Dim ImageSrc
		Dim Border
		Dim Width
		Dim Height
		Dim Alt
		Private Sub Class_Initialize()
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			
			ImageSrc  = ""	
			Width     = -1
			Border    = 0
			Height    = -1
			Alt = ""
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			ImageSrc = bag.Read("S")
			Border = Cint(bag.Read("B"))
			Width = Cint(bag.Read("W"))
			Height = Cint(bag.Read("H"))
			Alt = bag.Read("A")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "S",ImageSrc
			bag.Write "A",Alt
			bag.Write "B",Border
			bag.Write "W",Width
			bag.Write "H",Height
		End Function

	   
	   Public Function HandleClientEvent(e)
	   End Function			
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If

			 varStart = Now
			 			 			
			Response.Write "<IMG  id='" & Control.Name & "' " &_			
					   " SRC='" & ImageSrc & "'" &_
					   IIf(Alt<>""," Alt='" & Alt  + "' ","") &_
					   IIf(Width>0," Width=" & Width ,"") &_
					   IIf(Height>0," Height=" & Height,"") &_
					   " Border=" & Border &_
					   IIf(Control.Style<>""," Style='" & Control.Style  + "' ","") &_
					   IIf(Control.CssClass<>""," Class='" & Control.CssClass  + "' ","") &_
				     ">" & vbNewLine

			Page.TraceRender varStart,Now,Control.Name
		End Function

	End Class

%>