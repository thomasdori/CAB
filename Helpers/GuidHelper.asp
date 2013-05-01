<%
	Dim guid : Set guid = New GuidHelperClass

	Class GuidHelperClass
		Public Function GetGUID()
			Dim objTypeLib, guidNew

			Set objTypeLib = Server.CreateObject("ScriptLet.TypeLib")
			guidNew = Left(objTypeLib.GUID, 38)
			Set objTypeLib = Nothing

			GetGUID = guidNew
		End Function
	End Class
%>