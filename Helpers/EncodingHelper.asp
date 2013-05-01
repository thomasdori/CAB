<%
	'Dependencies: SanitationHelper.asp

	Dim encoder : Set encoder = New EncodingHelperClass

	Class EncodingHelperClass
		Private rangeMinimum
		Private rangeMax

		Private Sub Class_Initialize()
			Private rangeMinimum = 0
			Private rangeMax = 255
		End Sub

		Public Function Decode(text)
			Dim i

			If Nvl(text,"") <> "" Then
				For i = CONST_CHARACTER_RANGE_MIN to CONST_CHARACTER_RANGE_MAX
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
		' Just skips dot (46), slash (47) and colon (58) for date and time strings
		Public Function Encode(text)
			Dim i

			sanitation.RemoveUnallowedStrings(text)

			If Nvl(text,"") <> "" Then
				For i = CONST_CHARACTER_RANGE_MIN to CONST_CHARACTER_RANGE_MAX
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
		End Sub

		Private Function IsUnAllowedCharacter(value)
			IsUnAllowedCharacter = (value <> 46 And value <> 47 And value <> 58 And chr(value) <> " ")
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