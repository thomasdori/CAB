<%
	Class ProcedureParameter
		Dim key
		Dim value
		Dim type

		'A way to define a callable constructor with parameters
		Public Default Function Init(parameterKey, parameterValue, parameterType)
			key = "@" & parameterKey
			value = parameterValue
			type = parameterType

			Set Init = Me
		End Function

	End Class
%>