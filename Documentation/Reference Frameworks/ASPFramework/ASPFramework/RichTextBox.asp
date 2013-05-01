<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_RichTextBox.asp"    -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<!--Include File = "Home.asp"        -->
<%	Page.Execute %>	
<Span Class="Caption">RichTextBox Sample<br></Span>
<br>By Raymond Irving (xwisdom at GotDotNet)<br>
<STYLE TYPE="text/css">
a:hover {
	color: #FF0000;
}
.body {
	font-family: Verdana, Tahoma, Arial, 'MS Sans Serif';
	font-size: 77%;
}
.cbtn {
	BORDER-LEFT: #efedde 1px solid;
	BORDER-RIGHT: #efedde 1px solid;
	BORDER-TOP: #efedde 1px solid;
	BORDER-BOTTOM: #efedde 1px solid;
}
.txtbtn {font-family:tahoma; font-size:70%; color:menutext;}
FORM SELECT {
	font-family: 'Microsoft Sans Serif', Verdana, sans-serif;
	font-size: 8pt;
	color: #000000;
}
FORM INPUT {
	background-color: #efefef;
	font-family: 'Microsoft Sans Serif', sans-serif;
	font-size: 10pt;
}
</STYLE>
<%Page.OpenForm%>
<%cmdReadOnlyYN%> | <%cmdVisibleYN%> | <%cmdEnabledYN%> | <%cmdToolbarYN%><br><br>
<%txtRichBox%>
<%Page.CloseForm%>

</BODY>
</HTML>

<%  'This would normaly go in a another page, but for the sake of simplicity and to minimize the number of pages
    'I'm including code behind stuff here...
	'Response.AddHeader  "content-disposition", "attachment; filename=filename.ext"	
	
	Dim txtRichBox
	Dim cmdReadOnlyYN
	Dim cmdVisibleYN
	Dim cmdEnabledYN
	Dim cmdToolbarYN


	Public Function Page_Init()
 	
		Set cmdReadOnlyYN = New_ServerLinkButton("cmdReadOnlyYN")
		Set cmdVisibleYN = New_ServerLinkButton("cmdVisibleYN")
		Set cmdEnabledYN = New_ServerLinkButton("cmdEnabledYN")
		Set cmdToolbarYN = New_ServerLinkButton("cmdToolbarYN")

		Set txtRichBox = New_ServerRichTextBox("txtRichBox")
		
		'set image gallery url
		txtRichBox.ImageGalleryURL = "ImageGallery.asp"
		
	End Function

	Public Function Page_Controls_Init()

		txtRichBox.Width = 600
		txtRichBox.Height= 320
		txtRichBox.Text="<TABLE cellSpacing=1 cellPadding=2 width=""100%"" align=center border=0><TBODY><TR><TD  vAlign=top width=""100%"" bgColor=#ccff66><P align=center>&nbsp;<U>The <FONT size=4>Classic</FONT></U> <FONT size=7>ASP</FONT> <U><FONT size=4>Framework</FONT> <EM>is</EM> <FONT size=6>Awesome!</FONT></U></P></TD></TR></TBODY></TABLE><P></P><IMG src=""images/book01.gif"">"
		
		cmdReadOnlyYN.Text = "Toggle ReadOnly"
		cmdVisibleYN.Text  = "Toggle Visible"
		cmdEnabledYN.Text  = "Toggle Enabled"
		cmdToolbarYN.Text  = "Toggle Toolbar"
	End Function

	Public Function Page_Load()
		If txtRichBox.TextChanged Then
			Response.Write "<BR>Rich Text Box was modified<BR>"
			'Page.ShowMessage "HTML",txtRichBox.Text,100,50,200,150
		End If
	End Function

	Public Function cmdReadOnlyYN_OnClick()
		txtRichBox.ReadOnly = Not txtRichBox.ReadOnly
	End Function

	Public Function cmdVisibleYN_OnClick()		
		txtRichBox.Control.Visible = (Not txtRichBox.Control.Visible)
	End Function
	
	Public Function cmdEnabledYN_OnClick()
		txtRichBox.Control.Enabled = Not txtRichBox.Control.Enabled
	End Function

	Public Function cmdToolbarYN_OnClick()
		txtRichBox.HideToolbar = Not txtRichBox.HideToolbar
	End Function

%>