<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_CheckBox.asp" -->
<!--#Include File = "Controls\Server_DataRepeater.asp" -->
<!--#Include File = "Controls\Server_Label.asp"    -->
<!--#Include File = "DBWrapper.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Data Repeater Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">DataRepeater Sample</Span>

<%Page.OpenForm%>
	<%lblMessage%><HR>
	<HR>
	<%objRepeater%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	Dim lblMessage
	Dim objRepeater
	Dim objLabelMessage
	Dim chkEnablePagination
	
	Public Function Page_Init()
		Set objLabelMessage = New StringBuilder
		Set lblMessage = New_ServerLabel("lblMessage")
		Set objRepeater = New_ServerDataRepeater("objRepeater")
	End Function

	Public Function Page_Controls_Init()						
		lblMessage.Control.Style = "border:1px solid blue;background-color:#EEEEEE;width:100%;font-size:8pt"		
		objLabelMessage.Append   "This is an Example"
		objRepeater.Header = "fncHeader"
		objRepeater.ItemTemplate  = "fncItemTemplate"
		objRepeater.AlternateItemTemplate  = "fncAlternateItemTemplate"
		objRepeater.Footer = "fncFooter"

	End Function
	
	Public Function Page_PreRender()
		Set objRepeater.DataSource = GetRecordSet("Select top 20 CustomerID,CompanyName,ContactName + '/' + ContactTitle As Contact, Address From Customers")
		'objLabelMessage.Append "<B>NavigatorSize:</B>" & objRepeater.NavigatorSize & "<BR>"

		lblMessage.Text = objLabelMessage.ToString()
	End Function

	Public Function fncHeader()
		Response.Write "<SPAN STYLE='width:50%;color:white;background-color:#777777'>This is the header</span><BR>" 
	End Function
	
	Public Function fncFooter()
	Response.Write "<SPAN STYLE='width:50%;color:white;background-color:#777777'>This is the footer</span>"
	End Function

	Public Function fncItemTemplate(ds)
		Response.Write "<SPAN STYLE='width:50%;'>" & ds(0).Value & "</span><BR>"
	End Function

	Public Function fncAlternateItemTemplate(ds)
		Response.Write "<SPAN STYLE='width:50%;background-color:#DDDDDD'>" & ds(0).Value 
		If ds(0).Value = "BOTTM" or ds(0).Value = "DUMON" Then	
				Response.Write "  <input type=button value = 'Edit me!' "
				Response.Write Page.GetEventScript("onclick","Page","ButtonClicked",ds(0).Value,"")
				Response.Write ">"
		End If
		Response.Write "</span><BR>"
	End Function
	

	Public Function Page_ButtonClicked(e)
		objLabelMessage.Append  "You clicked the custom column.. value=" & e.Instance  & "<BR>"
	End Function

%>