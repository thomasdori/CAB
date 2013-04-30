<%
	Dim guid
	Set guid = new GuidHelper

	Class GuidHelper
		Function getGUID()
		    getGUID = CreateObject("Scriptlet.TypeLib").GUID
		End Function
	End Class
%>