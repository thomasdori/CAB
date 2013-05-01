<!--#Include File = "Controls/Server_AjaxProcessor.asp" -->
<%
	Sub Page_OnClientSideEvent(e)
		Dim txt
		txt = AjaxForm.GetFormValue("txtTest")
		Response.Write txt 'Eval(txt)
	End Sub
	
		

%>

