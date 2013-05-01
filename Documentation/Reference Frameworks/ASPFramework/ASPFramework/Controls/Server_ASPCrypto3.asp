<%
':: By John Hidey (jhidey at www.gotdotnet.com)
Class cASPCrypto
	Private Base64Chars
	
	Private Sub Class_Initialize()
		Base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	End Sub

	Public Function EncodeStr64(strIn)
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 3
			c1 = Asc(Mid(strIn, n, 1))
			c2 = Asc(Mid(strIn, n + 1, 1) + Chr(0))
			c3 = Asc(Mid(strIn, n + 2, 1) + Chr(0))
			w1 = Int(c1 / 4) : w2 = (c1 And 3) * 16 + Int(c2 / 16)
			If Len(strIn) >= n + 1 Then 
				w3 = (c2 And 15) * 4 + Int(c3 / 64) 
			Else 
				w3 = -1
			End If
			If Len(strIn) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode(w1) + mimeencode(w2) + mimeencode(w3) + mimeencode(w4)
		Next
		EncodeStr64 = strOut
	End Function

	Public Function DecodeStr64(strIn)
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 4
			w1 = mimedecode(Mid(strIn, n, 1))
			w2 = mimedecode(Mid(strIn, n + 1, 1))
			w3 = mimedecode(Mid(strIn, n + 2, 1))
			w4 = mimedecode(Mid(strIn, n + 3, 1))
			If w2 >= 0 Then _
				strOut = strOut + Chr(((w1 * 4 + Int(w2 / 16)) And 255))
			If w3 >= 0 Then _
				strOut = strOut + Chr(((w2 * 16 + Int(w3 / 4)) And 255))
			If w4 >= 0 Then _
				strOut = strOut + Chr(((w3 * 64 + w4) And 255))
		Next
		DecodeStr64 = strOut
	End Function

	Private Function mimeencode(intIn)
		If intIn >= 0 Then 
			mimeencode = Mid(Base64Chars, intIn + 1, 1) 
		Else 
			mimeencode = ""
		End If
	End Function

	Private Function mimedecode(strIn)
		If Len(strIn) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr(Base64Chars, strIn) - 1
		End If
	End Function
	
End Class
%>
