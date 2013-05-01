<!--#Include File = "Server_ASPViewState.asp" -->
<!--#Include File = "Server_ASPListItemCollection.asp" -->
<%
	Public Function New_ViewStateObject()
		Set New_ViewStateObject = New cASPViewState
	End Function

	Public Function New_ListItemsCollectionObject()
		Set New_ListItemsCollectionObject = new cASPListItemCollection	
	End Function
	
%>