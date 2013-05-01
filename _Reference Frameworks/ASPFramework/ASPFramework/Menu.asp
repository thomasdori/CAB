<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Menu.asp" -->
<!--#Include File = "Controls\Server_Label.asp" -->

<HTML>
<HEAD>
	<TITLE>Server Menu Example</TITLE>
	<LINK rel="stylesheet" type="text/css" href="Samples.css">
	<LINK rel="stylesheet" type="text/css" href="<%=SCRIPT_LIBRARY_PATH%>menu/item_style.css">
<script language='javascript'>

</script>
</HEAD>
<BODY id="CLASPBody" >
<Span Class="Caption">Menu Sample<hr></Span>
<%Page.Execute%>	
<!--Include File = "Home.asp"        -->
		<%lblMenuClicked%>
		<%mnMenu%>
<%Page.OpenForm%>
<%Page.CloseForm%>
</BODY>
</HTML>
<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim mnMenu
	Dim lblMenuClicked
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		Set mnMenu	= New_ServerMenu("mnMenu","")
		Set lblMenuClicked = New_ServerLabel("lblMenuClicked")
		mnMenu.Left =  10
		mnMenu.Top  =  100
	End Function

	Function Page_Controls_Init()	
		
	mnMenu.Add "Home 0","","",Array	( Array("Add Home","",""), _
									  Array("Add Right","",""), _
									  Array("Add Down","","") )

	mnMenu.Add "Home 1","","",Array	( Array("More...","",""), _
									  Array("Add Right","",""), _
									  Array("Add Down","","") )

			
	'mnMenu.Items(0).SetAttribute "CA","MyCustomAttr"
	End Function

	Function Page_Load()
	End Function
	
'WEBCONTROLS POSTBACK EVENT HANDLERS
	
	Function mnMenu_OnClick(m) 		
		lblMenuClicked.Text =  "Clicked on:" & m.Text
		Select Case m.Text
			Case "Add Home":
					mnMenu.Add "Home "  & mnMenu.Count ,"","",Array		( Array("Add Home","",""), _
											  Array("Add Right","",""), _
											  Array("Add Down","","") )
			Case "Add Right"
					m.Add "New Rigth"  & m.Count , "","",""
					m.Add "Add Right" , "","",""
					m.Add "Add Down" , "","",""

			Case "Add Down"
					m.Parent.Add "New Down" ,"","",Array		( Array("Add Home","",""), _
											  Array("Add Right","",""), _
											  Array("Add Down","","") )

		End Select
		
	End Function
%>
