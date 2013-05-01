<%

' ---------------------------------------------------------------
'  URL Helper
' ---------------------------------------------------------------
' 
' url_segmant
' url_array
' url_default
' url_seg_int
' do_url_segmant
' url_anchor_name




' ---------------------------------------------------------------
'  URL Segmants
' ---------------------------------------------------------------
' 
' 
' @string

function  url_segmant(num)
	
	url_segmant = do_url_segmant(num, "any")
	
end function




' ---------------------------------------------------------------
'  URL Array
' ---------------------------------------------------------------
' 
' 
' @string

function  url_array
	
	url_array = split( replace(CURRENTPATH, ".asp", "") , "\", -1, 0)
	
end function


' ---------------------------------------------------------------
'  URL Segmants for Intergers Only
' ---------------------------------------------------------------
' 
' 
' @string

function url_seg_int(num)

	url_seg_int = do_url_segmant(num, "int")

end function







' ---------------------------------------------------------------
'  Do URL Segmants
' ---------------------------------------------------------------
' 
' 
' @string

function do_url_segmant(num,chars)
	
	dim current_file, file_path_array, seg_return
	
	FILEPATH 		= replace(FILEPATH, "//","/")
	
	current_file 	= right(FILEPATH,len(FILEPATH) -1)
	
	file_path_array = Split(FILEPATH, "/", -1)
	
	for i = 1 to ubound(file_path_array)
	
		if num = (i) then
	
			seg_return =  file_path_array(num)
	
		end if
	
	next
	
	if chars = "any" then
		
		do_url_segmant = lcase( reg_replace(seg_return, ALLOWCHARS, "") )
	
	else
		
		do_url_segmant = is_number(seg_return)
	
	end if
	
	
end function



' ---------------------------------------------------------------
'  URL Segmants with Defualts
' ---------------------------------------------------------------
' 
' 
' @string

function url_default(num, val)
	
	if NOT does_exists( do_url_segmant(num, "any") ) then
		
		url_default = val
	
	else 
	
		url_default = do_url_segmant(num, "any")
	
	end if
	
end function



' ---------------------------------------------------------------
'  Convert Strings to URL Segmants
' ---------------------------------------------------------------
' 
' 
' @string

function url_anchor_name(var)
	
	url_anchor_name = lcase( reg_replace(var, "[^A-Za-z0-9]", "-") )
	
end function


%>