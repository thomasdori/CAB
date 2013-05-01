<!-- #include File="../Helpers/DateTimeHelper.asp" -->
<!-- #include File="../Libraries/adovbs.inc" -->

<!-- #include File="_Model.asp" -->

<%
	Dim errorModel : Set errorModel = new ErrorModelClass

	Class ErrorModelClass
		Function write(errorMessage)

			'todo: lock error to database

			'ASPCode: oError.ASPCode
			'ASPDescription: oError.ASPDescription
			'Category: oError.Category
			'Column: oError.Column
			'Description: oError.Description
			'File: oError.File
			'Line: oError.Line
			'Number: oError.Number
			'Source: oError.Source
			'Time: dateTime.GetCurrentDateTime()
			'User: ...
		End Function
	End Class
%>