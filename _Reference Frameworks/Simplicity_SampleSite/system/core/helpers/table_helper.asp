<%

' ---------------------------------------------------------------
'  Table Helpers
' ---------------------------------------------------------------
' 
' create_header
' table_row
' end_table
' table_no_data
' atl_row_style
' view_atl_row_style
' table_finish





' ---------------------------------------------------------------
'  Create Header Table
' ---------------------------------------------------------------
' 
' @var = string

function create_header(caption,header)
	
	dim local_html, add_class
	
	local_html = "<table id=""dataTable"" cellspacing=""0"">"
	
	if does_exists(caption) then
	
		local_html = local_html & "<caption>"&caption&"</caption>"
	
	end if
	
	local_html = local_html & "<tr>"
	
	for i = 0 to ubound(header)
		
		if i = 0 then
		
			add_class = "class=""nobg"""
		
		end if
		
		local_html = local_html & "<th "&add_class&" scope=""col"">" & header(i) & "</th>"
		
	next
	
	local_html = local_html & "</tr>"
	
	create_header  = local_html
	
end function




' ---------------------------------------------------------------
'  Create Header Table
' ---------------------------------------------------------------
' 
' @var = string

function table_row(row_ct,data)

	table_row = build_table_row(row_ct,data,"th")
	
end function 


function report_row(row_ct,data)
	
	report_row = build_table_row(row_ct,data,"td")
	
end function 




function build_table_row(row_ct,data,report)
	
	dim local_html
	
	local_html = "<tr>"
	
	for i = 0 to ubound(data)
	
		if not does_exists(data(i)) then
			data(i) = "&nbsp;"
		end if
		
		if i = 0 and report = "th" then
			
			local_html = local_html & "<td scope=""row""" & atl_row_style(row_ct,"th") & ">" & data(i) & "</th>" & vbCrLf
			
		else
			
			local_html = local_html & "<td" & atl_row_style(row_ct,"td") & ">" & data(i) & "</td>" & vbCrLf
		
		end if
	
	next
	
	local_html = local_html & "</tr>"
		
	build_table_row  = local_html
	
end function






' ---------------------------------------------------------------
'  End Table
' ---------------------------------------------------------------
' 
' @var = string
function table_no_data(colspan)
	
	table_no_data = "<tr><td colspan=""" & colspan & """><p><em>Sorry, this report returned no Data</em></p></td></tr>" & vbcrlf
	
end function



' ---------------------------------------------------------------
'  Create Alt Row Styles
' ---------------------------------------------------------------
' 
' @var = string

function atl_row_style(var,table_type)
	
	dim rowColor, row_ct
	
	row_ct = var mod 2
	
	if table_type = "th" then
	
		if row_ct <> 0 then
			
			rowColor = "specalt"
		else
			
			rowColor = "spec"
		
		end if
	
	else
	
		if row_ct <> 0 then
		
			rowColor = "alt"
		
		else
		
			rowColor = ""
		
		end if
	
	end if
	
	atl_row_style = " class="""&rowColor&""""
	
end function




' ---------------------------------------------------------------
'  view Alt Row Styles
' ---------------------------------------------------------------
' 
' @var = string

function view_atl_row_style(var,table_type)
	
	dim rowColor, row_ct
	
	row_ct = var mod 2
	
	if table_type = "th" then
	
		if row_ct <> 0 then
			
			rowColor = "specalt"
		else
			
			rowColor = "spec"
		
		end if
	
	else
	
		if row_ct <> 0 then
		
			rowColor = "alt"
		
		else
		
			rowColor = ""
		
		end if
	
	end if
	
	view_atl_row_style = rowColor
	
end function

' ---------------------------------------------------------------
'  End Table
' ---------------------------------------------------------------
' 
' @var = string

function end_table(ending)
	
	end_table = "</table>" & vbcrlf & ending & vbcrlf
	
end function





' ---------------------------------------------------------------
'  Finsih the TD's in a table
' ---------------------------------------------------------------
' 
' @var = string

function table_finish(count, total)
	
	dim local_str
	
	if count <> 1 then
	
		for i = count to total
			
			local_str = local_str & "<td>&nbsp;</td>"
		
		next
		
		local_str = local_str & "</tr>"
		
	end if
	
	table_finish = local_str
	
end function
%>