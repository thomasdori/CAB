<%
	Dim sanitation
	Set sanitation = new SanitationHelper

	Class SanitationHelper
		' Description: Function to remove unallowed strings like 'script'
		' Developer: tdor
		' Creation Date: 22.07.2012
		Public Function RemoveUnallowedStrings(value)
			if(value<>"") then
				value = Replace(value, "javascript", "")
		        value = Replace(value, "iframe", "")
				value = Replace(value, "script", "")
		        value = Replace(value, "<![CDATA[", "")
		        value = Replace(value, "]]>", "")
			end if

			RemoveUnallowedStrings = value
		End Function
	End Class
%>