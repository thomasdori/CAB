<%

    Class htmlTable
    
	public caption
	public cellpadding
	public table_style
	public td_style
	public headers
	public sizes
	public altRows
	public plocal
	public array_data
	public sql_conn
	public sql_stmt
	public array_type
	public data_type
	public sort_array
	public sort_segment
	dim local_start, local_table, local_headers, add_class, grabsize
	dim j, i, n, r, local_data, value, addstyle
	dim rowColor, LocalConn, LocalRs, data_check, add_align, local_str
	dim sort_var
	
	Private Sub Class_Initialize()
		sql_conn		= config_db
		caption 		= NULL
		headers 		= NULL
		table_style 	= NULL
		array_type 		= 2
		data_check 		= false
		altRows 		= true
		data_type		= NULL
		sort_array		= NULL
	End Sub
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	function build
		
		local_table = start_table(caption,cellpadding)
		
		local_table = local_table & create_headers(headers,sizes)
		
		if isArray(array_data) then
		
			local_table = local_table & create_array_data(array_data,td_style,array_type)
		
		else
		
			local_table = local_table & create_table_data(sql_stmt,td_style)
		
		end if
		
		local_table = local_table & end_table("")
		
		build = local_table
	
	end function
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	private function start_table(caption,cellpadding)
		
		if NOT does_exists(cellpadding) then
			cellpadding = 0
		end if 
		
		if does_exists(table_style) then
			
			local_start = "<table id="""&table_style&""" cellspacing="""&cellpadding&""">"
		
		else
				local_start = "<table id=""dataTable"" cellspacing="""&cellpadding&""">"
		
		end if
		
		if does_exists(caption) then
		
			local_start = local_start & "<caption>"&caption&"</caption>"
		
		end if
		
		start_table = local_start
		
	end function
	
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	
	private function create_headers(var,sizes)
		
		if isarray(var) then
			local_headers = "<tr>"
			
			for i = 0 to ubound(var)
				
				if i = 0 then
				
					add_class = "class=""nobg"""
				
				end if
				
				if isarray(sort_array) then
					
					sort_var = split(url_default(sort_segment, "XXXX:XXXX"), ":")
					
					if lcase(sort_var(0)) = lcase(sort_array(i)) then
						
						add_class = "class=""sort_active"""
					
					end if
				
				end if
				
				if isArray(sizes) then
					
					grabsize = " width="""&sizes(i)&""""
					
				end if
				
				local_headers = local_headers & "<th "&grabsize&" "&add_class&" scope=""col"">" & var(i) & "</th>"
				
				grabsize = ""
				
				add_class =""
				
			next
			
			local_headers = local_headers & "</tr>"
		end if
		
		create_headers  = local_headers
		
	end function
	
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	
	private function create_array_data(var,td_style,array_type)
		
		if array_type = 1 then
		
			for j = 0 to ubound(var)
			
				r = n Mod 2
				
				local_data = local_data & "<tr>"
				
				if isArray(td_style) then
					
					addstyle = "style="""&td_style(i)&""""
				
				end if
				
				value = var(j)
				
				local_data = local_data & "<td "&addstyle&" scope=""row"" "&alt_rows(r,"th")&">" & value & "</td>" & vbCrLf
				
				local_data = local_data & "</tr>"
				
				n=n+1
				
			next
		
		else
			
			for j = 0 to ubound(var, 1)
			
			r = n Mod 2
			
			local_data = local_data & "<tr>"
			
				for i = 0 to ubound(var, 2)
					
					if does_exists( var(j,i) ) then
						
						value = var(j,i)
					else
						
						value = "&nbsp;"
					
					end if
						
					if isArray(td_style) then
					
						addstyle = "style="""&td_style(i)&""""
					
					end if
					
					if i = 0 then
						
						local_data = local_data & "<td "&addstyle&" scope=""row"" "&alt_rows(r,"th")&">" & value & "</td>" & vbCrLf
						
					else
						
						local_data = local_data & "<td "&addstyle&" "&alt_rows(r,"td")&">" & value & "</td>" & vbCrLf
					
					end if
					
					value = ""
					
				next
			
			local_data = local_data & "</tr>"
			
			n=n+1
			
			next
			
		end if
		
		create_array_data  = local_data
			
	end function
	
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	
	private function create_table_data(sql_stmt,td_style)
		
		Set LocalConn = db_open(sql_conn)
		
		Set LocalRs = LocalConn.Execute(sql_stmt)
		
		n = 0
		
		do while not LocalRs.EOF
			
			data_check = "true"
			
			local_data = local_data & "<tr>"
			
			For i=0 To LocalRs.Fields.Count-1
				
				if does_exists( LocalRs(i) ) then
					
					value = LocalRs(i)
				else
					
					value = "&nbsp;"
				
				end if
				
				r = n Mod 2
				
				if isArray(td_style) then
				
					addstyle = "style="""&td_style(i)&""""
				
				end if
				
				if isarray(data_type) then
					add_align 	= get_data_type(data_type(i),"align", value)
					value		= get_data_type(data_type(i),"data", value)
				end if
				
				if i = 0 then
					
					local_data = local_data & "<td "&add_align&" "&addstyle&" scope=""row"" "&alt_rows(r,"th")&">" & value & "</td>" & vbCrLf
					
				else
					
					local_data = local_data & "<td "&add_align&" "&addstyle&" "&alt_rows(r,"td")&">" & value & "</td>" & vbCrLf
				
				end if
				
				value = ""
				
			next
			
			local_data = local_data & "</tr>"
			
			n = n + 1
		
		LocalRs.MoveNext
		
		Loop
		
		
		if data_check = false then
		
			local_data = table_no_data( ubound(headers) + 1 )
		
		end if
		
		create_table_data = local_data
		
		Call db_close(LocalRs)
		
		Call db_close(LocalConn)
		
	end function
	
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	
	private function get_data_type(data, ptype, value)
		
		if ptype = "align" then
			
			Select Case data
				Case "D"
					local_str = "align=""center"""
				Case "T"
					local_str = ""
				Case "M"
					local_str = "align=""right"""
				Case "N"
					local_str = "align=""right"""
			End Select
		
		else
		
			Select Case data
				Case "D"
					local_str = value
				Case "T"
					local_str = value
				Case "M"
					local_str = format_money(value)
				Case "N"
					local_str = format_number(value)
			End Select
		
		end if
		get_data_type = local_str
	end function
	
	
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	
	private function alt_rows(var,table_type)
		
		if table_type = "th" then
			
			if var <> 0 and altRows then
			
				rowColor = "alt"
			
			else
			
				rowColor = ""
			
			end if
		
		else
			
			if var <> 0 and altRows then
			
				rowColor = "alt"
			
			else
			
				rowColor = ""
			
			end if
		
		end if
		
		alt_rows = "class="""&rowColor&""""
		
	end function
	
	
End Class
%>