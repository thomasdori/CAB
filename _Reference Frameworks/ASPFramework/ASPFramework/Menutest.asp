<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Menu.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Server Menu Example</TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
<style>
.m0l0iout {
	font-family: Tahoma;
	font-size: 12px;
	text-decoration: none;
	padding: 4px;
	color: #FFFFFF;
}
.m0l0iover {
	font: 12px Tahoma;
	text-decoration: underline;
	padding: 4px;
	color: #FFFFFF;
}

/* level 0 outer */
.m0l0oout {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #4682B4;
}
.m0l0oover {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #5AA7E5;
}

/* level 1 inner */
.m0l1iout {
	font: 12px Tahoma;
	text-decoration: none;
	padding: 4px;
	color: #000000;
}
.m0l1iover {
	font: bold 12px Tahoma;
	text-decoration : none;
	padding: 4px;
	color: #000000;
}

/* level 1 outer */
.m0l1oout {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #46B446;
	filter: alpha(opacity=85);
}
.m0l1oover {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #5AE55A;
}

/* level 2 inner */
.m0l2iover {
	font: 12px Tahoma;
	text-decoration : none;
	padding: 4px;
	color: #000000;
}

/* level 2 outer */
.m0l2oout {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #B44646;
}
.m0l2oover {
	text-decoration : none;
	border : 1px solid #FFFFFF;
	background: #E55A5A;
}
</style>

</HEAD>

<BODY id="CLASPBody">
<script language="javascript">
/*
  --- menu level scope settins structure --- 
  note that this structure has changed its format since previous version.
  Now this structure has the same layout as Tigra Menu GOLD.
  Format description can be found in product documentation.
*/
var MENU_POS = [
{
	// item sizes
	'height': 24,
	'width': 120,
	// menu block offset from the origin:
	//	for root level origin is upper left corner of the page
	//	for other levels origin is upper left corner of parent item
	'block_top': 0,//99,
	'block_left': 0,//287,
	// offsets between items of the same level
	'top': 0,
	'left': 119,
	// time in milliseconds before menu is hidden after cursor has gone out
	// of any items
	'hide_delay': 200,
	'expd_delay': 200,
	'css' : {
		'outer': ['m0l0oout', 'm0l0oover'],
		'inner': ['m0l0iout', 'm0l0iover']
	}
},
{
	'height': 20,
	'width': 150,
	'block_top': 23,
	'block_left': 0,
	'top': 21,
	'left': 0,
	'css': {
		'outer' : ['m0l1oout', 'm0l1oover'],
		'inner' : ['m0l1iout', 'm0l1iover']
	}
},
{
	'block_top': 5,
	'block_left': 140,
	'css': {
		'outer': ['m0l2oout', 'm0l2oover'],
		'inner': ['m0l1iout', 'm0l2iover']
	}
}
]

</script>
<%Page.Execute%>	
<!--#Include File = "Home.asp"        -->
<%mnMenu%>
<%Page.OpenForm%>
<%Page.CloseForm%>
</BODY>
</HTML>

<%  '############ SERVER CODE

'VARIABLE DECLARATIONS
	Dim mnMenu
'PAGE EVENT HANDLERS	
	Function Page_Init()		
		Set mnMenu	= New_ServerMenu("mnMenu")
	End Function

	Function Page_Controls_Init()	
		
	mnMenu.Add "Home","","",Array		( Array("Add Home","","sb:test"), _
											  Array("Add Right","",""), _
											  Array("Add Down","","") )

	mnMenu.Add "Home 2","","",Array		( Array("More...","","sb:test"), _
											  Array("Add Right","",""), _
											  Array("Add Down","","") )

	End Function

	Function Page_Load()
	End Function
	
'WEBCONTROLS POSTBACK EVENT HANDLERS
	
	Function mnMenu_OnClick(m) 		
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
