<%

    Class Pagination
    
	public per_page
	public segmant
	public base_url
	
	public current_page
	public first_link
	public last_link
	public seperator
	public num_links
	public next_link
	public prev_link
	public sql
	public order_by
	public only_nextprev
	public order_type
	public sorting
	public start_order
	public end_order
	public div_class
	
	dim total_pages, build_links, i, intMinPages, order_1, order_2
	dim intMaxPages, SQLConn, LocalSql, LocalRs, rsArray, count, new_count, this_page
	dim LocalStr, SqlRegEx, add_sorting, add_end_order, conn
	
	Private Sub Class_Initialize()
		conn			= config_db
		per_page 		= 10
		segmant 		= 3
		base_url		= ""
		first_link 		= "&laquo; First"
		last_link		= "Last &raquo;"
		next_link		= "&gt;"
		prev_link		= "&lt;"
		seperator		= "&nbsp;"
		current_page 	= 1
		num_links		= 3
		only_nextprev	= FALSE
		order_type		= "ASC"
		div_class		= "pagination"
		
	End Sub

'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------


	function create_sql
		
		if does_exists(url_seg_int(segmant)) then
		
			current_page = url_seg_int(segmant)
		
		end if
		
		if count < per_page * current_page then
		
			new_count = count - (per_page * (current_page - 1))
		
		else
		
			new_count = per_page
		
		end if
		
		
		order_1 = order_type
		
		order_2 = get_order_type_dif(order_type)
		
		if does_exists(end_order) then
			
			add_end_order = ", " & end_order
		
		end if
		
		LocalStr = "select * from ( select top " & new_count & " * from ( "
		
		
		Set SqlRegEx = New RegExp
			
			SqlRegEx.Pattern = "SELECT"
			
			SqlRegEx.IgnoreCase = True
			
			SqlRegEx.Global = True
			
			LocalStr = LocalStr & SqlRegEx.Replace(sql, "select top " & per_page * current_page)
 		
		Set SqlRegEx = nothing
		
		LocalStr = LocalStr & " order by " & order_by & " " & order_1 & add_end_order
		
		LocalStr = LocalStr & " ) as newtbl order by " & order_by & " " & order_2 & add_end_order & " ) as newtbl2 order by " & order_by & " " & order_1 & add_end_order
		
		create_sql = LocalStr
		
	end function
	
	
	function create_links
		
		if does_exists(sorting) then
		
			add_sorting = "/" & sorting
		
		end if
		
		if does_exists(url_seg_int(segmant)) then
			
			current_page = url_seg_int(segmant)
			
		end if
		
		if NOT only_nextprev then
		
			count = getCount
			
			if count > per_page then
				
				total_pages = round(count/per_page)
				
				if total_pages * per_page < count then
				
					total_pages = total_pages + 1
				
				else
				
					total_pages = total_pages
					
				end if
				
				if (current_page > num_links) then
				
					intMinPages = current_page - num_links
				
				else
			
					intMinPages = num_links
				
				end if
					
				for i = intMinPages to current_page
				
					if int(current_page) <> i then
				
						if (i >= 1) then
				
							build_links = build_links & anchor(base_url & "/" & i & add_sorting,i,"") & seperator
				
						end if
				
					end if
				
				next
				
				build_links = build_links & "<strong>" & seperator & current_page & seperator & "</strong>" & seperator
				
				if (int(current_page) + num_links < total_pages) then
				
					intMaxPages = current_page + num_links
				
				else
				
					intMaxPages = total_pages
				
				end if
				
				for i = current_page to intMaxPages
				
					if int(current_page) <> i then
				
						if (i >= 1) then
				
							build_links = build_links & anchor(base_url & "/" & i & add_sorting,i,"") & seperator
				
						end if
				
					end if
				
				next
				
				if int(current_page) <> 1 then
				
					build_links = anchor(base_url & "/" & int(current_page) - 1 & add_sorting,prev_link,"") & seperator & build_links
					
					build_links = anchor(base_url & "/1" & add_sorting,first_link,"") & seperator & build_links
					
				end if
				
				if int(current_page) <> total_pages then
					
					build_links = build_links & anchor(base_url & "/" & int(current_page) + 1 & add_sorting,next_link,"")
					
					build_links = build_links & seperator & anchor(base_url & "/" & total_pages & add_sorting,last_link,"")
					
				end if
			
				create_links =  build_links
			
			end if
		
		else
		
			count = getCount
			
			if count > per_page then
				
				total_pages = round(count/per_page)
				
				if total_pages * per_page < count then
				
					total_pages = total_pages + 1
				
				else
				
					total_pages = total_pages
					
				end if
				
				if int(current_page) = 1 then
					
					'build_links = build_links & prev_link & seperator
				
				else
				
					build_links = build_links & anchor(base_url & "/" & int(current_page) - 1 & add_sorting,prev_link,"") & seperator
					
				end if
				
				if int(current_page) = total_pages then
					
					'build_links = build_links & next_link
					
				else
					
					build_links = build_links & anchor(base_url & "/" & int(current_page) + 1 & add_sorting,next_link,"")
				
				end if
				
			end if
			
		end if
		
		
		build_links = "<div class=""" & div_class & """>" & build_links &"</div>"
		
		create_links =  build_links
		
	end function
	
	
	function get_order_type_dif(var)
	
		if ucase(var) = "ASC" then
		
			get_order_type_dif = "DESC"
			
		else
			get_order_type_dif = "ASC"
			
		end if
	
	end function
	
	
	
	
	function getCount
		
		Set SQLConn = db_open(conn)
		
		Set LocalRs = SQLConn.Execute(sql)
			
			if NOT LocalRs.EOF then
				
				rsArray = LocalRs.GetRows() 
        		
				getCount = UBound(rsArray, 2) + 1 
				
			else
			
				getCount = 0
			end if
			
		Call db_close(SQLConn)
		
	end function
	
end class
%>