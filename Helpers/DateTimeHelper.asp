<%
	Dim dateTime : Set dateTime = new DateTimeHelperClass

	Class DateTimeHelperClass
		Function GetCurrentDateTime()
			GetCurrentDateTime = Now
		End Function
	End Class
%>