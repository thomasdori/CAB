<%
	Dim guid
	Set guid = New GuidHelper

	Class GuidHelper
		Public Function GetGUID()
		    Set objTypeLib = Server.CreateObject("ScriptLet.TypeLib")
			guidNew = Left(objTypeLib.GUID, 38)
			Set objTypeLib = Nothing

			GetGUID = guidNew
		End Function
	End Class
%>