<%

' ---------------------------------------------------------------
'  Icon Helper
' ---------------------------------------------------------------
' 
' show_icon


' ---------------------------------------------------------------
'  Show Icon Standard
' ---------------------------------------------------------------
' 
'  @string

function show_icon(icon)
	
	show_icon = build_icon(icon & ".png", "absmiddle", "")
	
end function





' ---------------------------------------------------------------
'  Build Standard Icon
' ---------------------------------------------------------------
' 
'  @array

function build_icon(icon, align, attributes)
	
	dim atts
	
	atts	= add_attributes(attributes)
	
	build_icon = "<img src=""/public/icons/"&icon&""" width=""16"" height=""16"" alt="""" border=""0"" align="""&align&""" "&atts&" />"
	
end function
%>