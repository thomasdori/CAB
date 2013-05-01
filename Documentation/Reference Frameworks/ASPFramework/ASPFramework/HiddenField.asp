<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_HiddenField.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Hiden Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<script language="javascript">
	function changeTxtID() {
		var t = prompt("Enter value for txtID",clasp.getObject("txtID").value);
		if(t!="") {
			clasp.getObject("txtID").value = t;
			clasp.form.doPostBack("DoNothing","Page");
		}
	}
</script>
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%
	Call Main()
%>	
<Span Class="Caption">Hidden Fields... (view the source to see the hidden fields)<hr></Span>
<%Page.OpenForm%>
<a href="#" onclick="changeTxtID()">Change txtID hidden field...</a>
<%txtID%>
<%txtAnotherHiddenField%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim txtID
	Dim txtAnotherHiddenField
	
	Public Function Page_Init()
		Set txtID = New_ServerHiddenField("txtID")
		Set txtAnotherHiddenField = New_ServerHiddenField("txtAnotherHiddenField")		
		
		txtID.RaiseOnChanged = True
	End Function

	Public Function Page_Controls_Init()						
		txtID.Value = "A value"
		txtAnotherHiddenField.Value = "Another value"
	End Function
	
	Public Function Page_LoadViewState()
	End Function
	
	Public Function txtID_OnChanged(e,p)
		Response.Write "The value of txtID changed to " & txtID.Value
	End Function
%>