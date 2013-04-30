<%
	'Dependencies: SanitationHelper.asp

	Dim encoder
	Set encoder = new EncodingHelper

	Class EncodingHelper
		Public Function Decode(sText)
			if Nvl(stext,"") <> "" then
				Dim i
				For i=0 to 127
					if (i<>58 and i<>46 and i<>35  and i<>59 and chr(i)<>" ") then stext= Replace(stext, "&#" & i & ";", chr(i))
					if i=47 then i=58  'skip 0-9
					if i=64 then i=91  'skip a-z
					if i=96 then i=123 'skip A-Z
				next
			end if
			Decode= sanitation.RemoveUnallowedStrings(stext)
		End Function

		'58 - :
		'46 - .
		'45 - -
		'35 - #
		'38 - &
		'59 - ;
		Public Function Encode(text)
			if Nvl(text,"") <> "" then
				Dim i
				For i=0 to 127
					if (i<>58 and i<>45 and i<>46 and i<>47 and i<>35 and i<>38 and i<>59 and chr(i) <> " ") then
						text = Replace(text, chr(i), "&#" & i & ";")
					end if

					if i=47 then i=58  'skip 0-9
					if i=64 then i=91  'skip a-z
					if i=96 then i=123 'skip A-Z
				next
			end if
			encode= text
		End Function

		Private Function Nvl(feld, nullwert)
		   if Isnull(feld) then
		      Nvl = nullwert
		   elseif feld="" then
		      Nvl=nullwert
		   else
		      Nvl = feld
		   end if
		End Function
	End Class
%>