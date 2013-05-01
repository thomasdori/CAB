<!-- #include File="_Model.asp" -->

<%
	Dim userModel : Set userModel = New UserModelClass

	Class UserModelClass
		Public Function GetUser(userId)
			Dim parameterList : Set parameterList = New List
			parameterList.Add((New ProcedureParameter)("userId", userId, adInteger))

			model.CallProcedure("sp_getUser", parameterList, GetRef(HasContentHandler), GetRef(RowHandler))
		End Function

		Private Sub HasContentHandler(recordset)
			'todo: implement
		End Sub

		Private Sub RowHandler(recordset)
			'todo: implement
		End Sub
	End Class
%>