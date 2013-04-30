<%
	Dim guid
	Set guid = new GuidHelper

	Class GuidHelper
		Function getGUID()
		    Set objTypeLib = Server.CreateObject("ScriptLet.TypeLib")
			guidNew = Left(objTypeLib.GUID, 38)
			Set objTypeLib = Nothing

			getGUID = guidNew
		End Function
	End Class
%>