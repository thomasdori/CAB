<%

dim FILEPATH
dim current_file_path, current_file, file_path_array, file_check
dim i, build_file_path, Seperator, fso
dim new_file_path, objFile, strSource, footer_links
dim iSpace, site_content, content_title, content_main
dim CURRENTPATH, fatal_error

' ---------------------------------------------------------------
' Remove common extenstion
' ---------------------------------------------------------------
FILEPATH		= build_filepath
FILEPATH		= Replace(FILEPATH,".xml","")
FILEPATH		= Replace(FILEPATH,".asp","")
CURRENTPATH		= FILEPATH
fatal_error 	= false

if NOT reg_search(FILEPATH, ALLOWCHARS) = false then
	
	SystemInclude("system\application\errors\URL_error.asp")
	
	fatal_error = true
	
end if


' ---------------------------------------------------------------
' Create Controller from FILEPATH
' ---------------------------------------------------------------

if fatal_error = false then
	
	current_file_path 	= Server.MapPath("system\application\controller\")
	current_file 		= right(FILEPATH,len(FILEPATH) -1)
	file_path_array 	= Split(current_file, "/")
	file_check = "No"
	
	
	
	
	for i = 0 to ubound(file_path_array)
		if does_exists(file_path_array(i)) then
			seperator = "\"
		end if
		
		build_file_path = build_file_path & seperator & file_path_array(i)
		
		set fso = createobject("scripting.filesystemobject")
		
		CURRENTPATH = build_file_path & EXT
		
		if fso.FileExists (current_file_path & CURRENTPATH) then 
			SystemInclude("system\application\controller" & CURRENTPATH)
			file_check 	= "Yes"
			if file_check = "Yes" then exit for
		end if
	next
	
	build_file_path = ""
	seperator = ""
	
	
	if does_exists(current_file) Then
		if file_check <> "Yes" then
			for i = 0 to ubound(file_path_array)
				if i <> 2 then 
					seperator = "\"
				end if
				
				build_file_path = build_file_path & seperator & file_path_array(i)
				
				set fso = createobject("scripting.filesystemobject")
				
				
				if fso.FileExists (current_file_path & build_file_path & "\index" & EXT) then
					CURRENTPATH = build_file_path & "\index" & EXT
				end if
			next
			
			if fso.FileExists (current_file_path & CURRENTPATH) then 
				
				SystemInclude("system\application\controller" & CURRENTPATH)
				
				file_check 	= "Yes"
			
			end if
		end if
	end if
	
	' ---------------------------------------------------------------
	' 	this Checks the src url to drop
	'	the cookie and redirect
	' ---------------------------------------------------------------
	
	if lcase(url_segmant(1)) = "src" then
		
		Call cookie_write("src",url_segmant(2))
		
		Call redirect("/", "200")
		
	end if
	
	
	' ---------------------------------------------------------------
	' 	If nothing is loaded it loads the index page
	' ---------------------------------------------------------------
	
	if NOT does_exists(current_file) Then
		
		CURRENTPATH = "/index.asp"
		
		SystemInclude("system\application\controller\index.asp")
		
		file_check = "Yes"
	end if
	
	' ---------------------------------------------------------------
	' 	Once last check, then loads the 404 page
	' ---------------------------------------------------------------
	
	
	if file_check = "No" then
		
		SystemInclude("system\application\errors\404_error.asp")
	
	end if
	
end if 'End fatal_error if


function build_filepath
	
	dim local_str
	
	local_str = sv_get_value("QUERY_STRING")
	
	local_str = replace(local_str, "404;", "")
	
	local_str = replace(local_str, config_site_url, "")
	
	local_str = replace(local_str, PAGEEXT, "")
	
	local_str = "/" & local_str
	
	if not does_exists(local_str) then
		
		local_str = "/index"
		
	end if
	
	build_filepath = local_str 
	
end function
%>