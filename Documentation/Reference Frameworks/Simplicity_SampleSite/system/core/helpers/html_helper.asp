<%

' ---------------------------------------------------------------
'  Html Helpers
' ---------------------------------------------------------------
' 
' anchor
' anchor_popup
' mailto 
' remove_html
' remove_tags
' nb
' h4
' h3
' h2
' h1
' ul



' ---------------------------------------------------------------
'  Anchor
' ---------------------------------------------------------------
' 
' @url = string
' @name = string
' @attributes = array

function anchor(url,name, attributes)
	
	dim pLocal, atts
	
	if left(url,1) = "/" then
		
		url = right( url,len(url) - 1 )
	
	end if
	
	atts	= add_attributes(attributes)
	
	if not does_exists(url) then
	
		url = "index"
	
	end if
	
	if instr(url, "http") <> 0 then
		
		pLocal = "<a href="""&url&""" title="""& remove_html(name) &""" "&atts&">"&name&"</a>"
		
	elseif instr(url, "#") <> 0 then
		
		pLocal = "<a href="""&url&""" title="""& remove_html(name) &""" "&atts&">"&name&"</a>"
		
	else
		
		url = lcase(url) & PAGEEXT
		
		pLocal = "<a href=""/"&url&""" title="""& remove_html(name) &""" "&atts&">"&name&"</a>"
	
	end if
	
	anchor = pLocal

end function



' ---------------------------------------------------------------
'  Mailto
' ---------------------------------------------------------------
' 
' 
' @url = string
' @name = string
' @attributes = array

function anchor_popup(url, name, attributes, pop_attributes)
	
	dim local_str, local_url, atts, pop_atts
	
	if left(url,1) = "/" then
		
		url = right( url,len(url) - 1 )
	
	end if
	
	if instr(url, "http") <> 0 then
		
		local_url = url
	
	elseif instr(url, "#") <> 0 then
	
		local_url = url
	
	else
	
		local_url = "/"&url
	
	end if
	
	atts	= add_attributes(attributes)
	
	if does_exists(pop_attributes) then
	
		pop_atts = add_pop_attributes(pop_attributes)
	
	else
	
		pop_atts = "width=600,height=500,scrollbars=yes,status=yes,resizable=yes,screenx=50%,screeny=50%"
	
	end if
	
	local_str = "<a href='javascript:void(0);' " &_
		"onclick=""window.open('" & local_url & "', '_blank', " &_
		"'" & pop_atts & "');"" " &_
		"title=""" & remove_html(name) & """ " & atts & ">" & remove_html(name) & "</a>"
	
	anchor_popup = local_str
	
end function



' ---------------------------------------------------------------
'  Mailto
' ---------------------------------------------------------------
' 
' @email = string
' @name = string

function mailto(email,name)
	
	dim pLocal
	
	pLocal = "<a href=""mailto:"&email&""" title=""Email: "&email&""">"&name&"</a>"
	
	mailto = pLocal

end function




' ---------------------------------------------------------------
'  Remove Html
' ---------------------------------------------------------------
' 
' @var = string
' 

function remove_html(var)
	
	remove_html = reg_replace(var, "<[^>]*>", "")
	
End Function




' ---------------------------------------------------------------
'  Remove Tags
' ---------------------------------------------------------------
' 
' @var = string
' @tags = string

function remove_tags(var, tags)
	
	remove_tags = reg_replace(var, "<" & tags & "[^>]*>", "")
	
End Function





' ---------------------------------------------------------------
'  Add Non Breaking Spaces
' ---------------------------------------------------------------
' 
' @var = string
' 

function nb(var)
	
	Dim local_html, i
	
	for i=0 to var
		
		local_html = local_html & "&nbsp;"
		
	next
	
	nb = local_html
	
end function




' ---------------------------------------------------------------
'  H4 Html Tag
' ---------------------------------------------------------------
' 
' @var = string
' 

function h4(var)
	
	h4 = create_title(var,4)
	
end function




' ---------------------------------------------------------------
'  H3 Html Tag
' ---------------------------------------------------------------
' 
' @var = string
' 

function h3(var)
	
	h3 = create_title(var,3)
	
end function




' ---------------------------------------------------------------
'  H2 Html Tag
' ---------------------------------------------------------------
' 
' @var = string
' 

function h2(var)
	
	h2 = create_title(var,2)
	
end function




' ---------------------------------------------------------------
'  H1 Html Tag
' ---------------------------------------------------------------
' 
' @var = string
' 

function h1(var)
	
	h1 = create_title(var,1)
	
end function




' ---------------------------------------------------------------
'  Build Title Tag
' ---------------------------------------------------------------
' 
' @var = string
' @num = int

function create_title(var, num)
	
	create_title = "<h"&num&">" & widont(var) & "</h"&num&">" & Vbcrlf
	
end function




' ---------------------------------------------------------------
'  Update Text Helper for WYSIWYG
' ---------------------------------------------------------------
' 
' @var = string
' 

function update_textarea(var)
	
	dim local_str
	
	local_str		= remove_tags(var,"span")
	
	local_str		= replace(local_str, "&nbsp;", " ")
	
	update_textarea = local_str
	
end function



' ---------------------------------------------------------------
'  Create HTML List from an Array
' ---------------------------------------------------------------
' 
' @var = string
' 

function ul(var)
	
	dim items
	
	for i = 0 to ubound(var)
		
		items = items & "<li>" & var(i) & "</li>" & vbcrlf
	
	next
	
	ul = "<ul>" & vbcrlf & items & "</ul>" & vbcrlf
	
end function

%>