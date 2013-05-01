<%

' ---------------------------------------------------------------
'  Captcha Helper
' ---------------------------------------------------------------
' 
' captcha


' ---------------------------------------------------------------
'  Set Captcha Image and Session
' ---------------------------------------------------------------
' 
'  @string

function captcha (color_id, val)
	
	Dim objCanvas, new_text, new_width, new_x
	
	new_text = UCASE(val)
	
	Call session_write("captcha_session", Replace(new_text, " ","") )
	
	Set objCanvas = New Canvas
	
	Select Case color_id
	
		Case 1 'Blue Defualt
			objCanvas.GlobalColourTable(0) = RGB(51,102,153)
			objCanvas.GlobalColourTable(1) = RGB(255,255,255)
		
		Case 2 'Dark Blue
			objCanvas.GlobalColourTable(0) = RGB(23,49,136)
			objCanvas.GlobalColourTable(1) = RGB(255,255,255)
		
		Case 2 'Orange
			objCanvas.GlobalColourTable(0) = RGB(195,77,10)
			objCanvas.GlobalColourTable(1) = RGB(255,255,255)
			
		Case else
			objCanvas.GlobalColourTable(0) = RGB(51,102,153)
			objCanvas.GlobalColourTable(1) = RGB(255,255,255)
		
	End Select
	
	objCanvas.BackgroundColourIndex = 0
	
	objCanvas.ForegroundColourIndex = 1
	
	objCanvas.Resize 100,24, False
	
	new_width = objCanvas.GetTextWEWidth (new_text)
	
	new_x = (100-new_width) / 2
	
	new_x = Round(new_x)
	
	objCanvas.DrawTextWE new_x, 5, new_text
	
	objCanvas.Write

end function
%>