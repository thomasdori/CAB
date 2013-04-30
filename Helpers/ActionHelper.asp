<%
	Dim action : Set action = New ActionHelperClass

	Class ActionHelperClass
		Public Function Execute(eventName)
			'Dim fnc
			'Execute = False
'
			'Set fnc = GetFunctionReference(eventName)
			'If Not fnc Is Nothing Then
			'	Execute = True
			'	Call fnc()
			'End If
		End Function

		Public Function GetFunctionReference(sFncName)
			'On Error Resume Next
			''	Set GetFunctionReference = Nothing
			''	Set GetFunctionReference = GetRef(sFncName)
			'On error Goto 0
		End Function
	End Class
%>