<%

    Class breadcrumbs
    
	public seperator
	public spacer
	public home_show
	public home_link
	public crumbs
	public started_crumbs
	public skip_first
	dim build_array, segs, local_html, indx, url, this_space, array_count
	
	Private Sub Class_Initialize()
		spacer			= " "
		seperator 		= "&raquo;"
		home_show 		= true
		home_link 		= "/"
	End Sub
	
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
	
	
	function build
		
		this_space = spacer & seperator & spacer
		
		if home_show = true then
			
			local_html = anchor("", "Home", "")
			
			local_html = local_html & this_space
			
		end if
		
		if isarray(started_crumbs) then
		
			For indx=0 to ubound(started_crumbs)
				
				local_html = local_html & anchor(started_crumbs(indx+1), humanize_url( started_crumbs(indx) ), "") & this_space
				indx = indx + 1
			next
		
		end if
		
		if isarray(crumbs) then
			
			build_array = crumbs
			
			For indx=0 to ubound(build_array)
				
				if indx+1 = ubound(build_array) then
					
					local_html = local_html & humanize_url( build_array(indx) ) & this_space
					
				else
					
					local_html = local_html & anchor(build_array(indx+1), humanize_url( build_array(indx) ), "") & this_space
				
				end if
				
				indx = indx + 1
				
			next
			
		else
			
			build_array = url_array
			
			if build_array( ubound(build_array) ) = "index" then
				
				array_count = ubound(build_array) - 1
				
			else
				
				array_count = ubound(build_array)
				
			end if
			
			
			For indx=1 to array_count
				
				url = url & "/" & build_array(indx)
				
				if indx = array_count then
					
					local_html = local_html & humanize_url( build_array(indx) ) & this_space
					
				else
					
					if NOT skip_first and NOT indx = 0 then
						
						local_html = local_html & anchor(url, humanize_url( build_array(indx) ), "") & this_space
					
					end if
					
				end if
				
			next
			
		end if
		
		
		local_html = left(local_html, len(local_html) - len(spacer+seperator+spacer))
		
		build = local_html
		
	end function
	
End Class
%>