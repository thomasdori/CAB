<%
' ---------------------------------------------------------------
'	Custom Helpers per this Application
' ---------------------------------------------------------------
' 
' custom_helper
' top_nav
' 

'---------------------------------------------------------------------
' custom helper
'
' @var = string

function custom_helper(var)
	
	custom_helper = "<div class=""yellow_box"">" & var & "</div>"
	
end function

'---------------------------------------------------------------------



'---------------------------------------------------------------------
' Top Nav helper
'
' @var = string

function top_nav
	
	dim lists
	
	lists = array(_
				anchor("/", "Home", "")_
				, anchor("about-us", "About Us", "")_
				, anchor("getting-started", "Getting Started", "")_
				, anchor("sample-form", "Sample Form", "")_
				)
	
	top_nav = ul(lists)
	
end function

'---------------------------------------------------------------------




%>