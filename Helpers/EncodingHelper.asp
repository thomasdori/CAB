<%
	'Dependencies: SanitationHelper.asp

	Dim encoder
	Set encoder = new EncodingHelper

	Class EncodingHelper
		Function decode(sText)
			if nvl(stext,"") <> "" then
				Dim i
				For i=0 to 127
					if (i<>58 and i<>46 and i<>35  and i<>59 and chr(i)<>" ") then stext= Replace(stext, "&#" & i & ";", chr(i))
					if i=47 then i=58  'skip 0-9
					if i=64 then i=91  'skip a-z
					if i=96 then i=123 'skip A-Z
				next
			end if
			decode= sanitation.removeUnallowedStrings(stext)
		End Function

		'58 - :
		'46 - .
		'45 - -
		'35 - #
		'38 - &
		'59 - ;
		Function encode(text)
			if nvl(text,"") <> "" then
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

		Function nvl(feld, nullwert)
		   if isnull(feld) then
		      nvl = nullwert
		   elseif feld="" then
		      nvl=nullwert
		   else
		      nvl = feld
		   end if
		End Function
	End Class
%>