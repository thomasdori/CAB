<%

' ---------------------------------------------------------------
'  Directory Helper
' ---------------------------------------------------------------
' 
' dir_get_files
' dir_map_url
' dir_get_folder_list



' ---------------------------------------------------------------
'  Directory Listing
' ---------------------------------------------------------------
' 
' 
' 

function dir_get_files(path)
	
	dim local_str, fs, folder, file, item, url, temp
	
	set fs = CreateObject("Scripting.FileSystemObject")
	
	set folder = fs.GetFolder(path)
	
	local_str = local_str & "<ul>" & vbCrLf
	
	local_str = local_str & "<li><b>" & humanize(folder.Name) & "</b>" & Vbcrlf
	
	for each item in folder.SubFolders
		
		temp = dir_get_files(item.Path)
		
		local_str = local_str & temp
	
	next
	
    for each item in folder.Files
	
		url = dir_map_url(item.path)
		
		' item.Size
		' item.DateLastModified
		
		local_str = local_str & "<ul>" & vbCrLf
		
		if item.Name <> "index.asp" then
			
			local_str = local_str & "<li>" & anchor(url, Replace(humanize(item.Name), ".asp", ""), "") & Vbcrlf
			
		end if
		
		local_str = local_str & "</ul>" & vbCrLf
		
	next
	
	local_str = local_str & "</ul>" & vbCrLf
	
	dir_get_files = local_str
	
end function


' ---------------------------------------------------------------
'  Directory URL Listing
' ---------------------------------------------------------------
' 
' 
' 

function dir_map_url(path)
	
	dim rootPath, url, local_str
	
	rootPath = Server.MapPath("/")
    
	url = Right(path, Len(path) - Len(rootPath))
    
	local_str	= url
	
	local_str 	= Replace(local_str, "\", "/")
	
	local_str 	= Replace(local_str, "/system/application/controller/", "/")
	
	dir_map_url = Replace(local_str, ".asp", "")
	
end function


' ---------------------------------------------------------------
'  Directory Listing of a folder
' ---------------------------------------------------------------
' 
' 
' 

function dir_get_folder_list(path)
	
	dim local_str, fs, folder, file, item, url, link_name, temp
	
	set fs = CreateObject("Scripting.FileSystemObject")
	
	set folder = fs.GetFolder(path)
	
	for each item in folder.Files
	
		url = dir_map_url(item.path)
		
		if item.Name <> "index.asp" then
			
			link_name = item.Name
			link_name = Replace(link_name, EXT, "")
			link_name = humanize_url(link_name)
			link_name = trim(link_name)
			
			local_str = local_str & "<li>" & anchor(url, link_name, "") & Vbcrlf
			
		end if
		
	next
	
	dir_get_folder_list = "<ul>" & local_str & "</ul>"
	
end function


%>