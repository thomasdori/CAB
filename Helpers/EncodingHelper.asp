<!-- #include File="SanitationHelper.asp" -->

<%
	'Dependencies: SanitationHelper.asp

	Dim encoder : Set encoder = New EncodingHelperClass

	Class EncodingHelperClass
		Private rangeMin
		Private rangeMax

		Private Sub Class_Initialize()
			rangeMin = 0
			rangeMax = 127
		End Sub

		Public Function Decode(text)
			Dim i

			If Nvl(text,"") <> "" Then
				For i = rangeMin To rangeMax
					If (IsUnAllowedCharacter(i)) Then
						text = Replace(text, "&#" & value & ";", chr(i))
					End If

					i = SkipAlphaNumerics(i)
				Next
			End If
			Decode = sanitation.RemoveUnallowedStrings(text)
		End Function

		' Basically like Server.HTMLEncode
		' But encodes even more
		Public Function Encode(text)
			Dim i

			sanitation.RemoveUnallowedStrings(text)

			If Nvl(text,"") <> "" Then
				For i = rangeMin To rangeMax
					If (IsUnAllowedCharacter(i)) Then
						text = Replace(text, chr(i), "&#" & i & ";")
					End If

					i = SkipAlphaNumerics(i)
				Next
			End If
			Encode = text
		End Function


		Public Function EncodeUrl(text)
			EncodeUrl = Server.URLEncode(text)
		End Function

		'10 - CR
		'13 - LF
		'32 - SPACE
		'35 - 35
		'38 - &
		'46 - .
		'47 - /
		'58 - :
		'59 - ;
		Private Function IsUnAllowedCharacter(i)
			' Just skips dot (46), slash (47) and colon (58) for date and time strings
			IsUnAllowedCharacter = (i<>10 And i <> 13 And i<>32 And i<>35 And i<>38 And i<>46 And i<>47 And i<>58 And i<>59)
		End Function

		Private Function SkipAlphaNumerics(value)
			If value = 47 Then value = 58  'skip 0-9
			If value = 64 Then value = 91  'skip a-z
			If value = 96 Then value = 123 'skip A-Z

			SkipAlphaNumerics = value
		End Function

		Private Function Nvl(feld, nullwert)
		   If Isnull(feld) Then
		      Nvl = nullwert
		   Elseif feld = "" Then
		      Nvl = nullwert
		   Else
		      Nvl = feld
		   End If
		End Function
	End Class
%>