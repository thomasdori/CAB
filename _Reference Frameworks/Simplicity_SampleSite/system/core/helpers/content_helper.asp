<%

' ---------------------------------------------------------------
'  Content Helper
' ---------------------------------------------------------------
' 
' load_content
' content_file
' content_parse_asp





' ---------------------------------------------------------------
'  Load External Content
' ---------------------------------------------------------------
' 
'  @string

function content_load
	
	dim file_path, file_name, local_str
	
	file_path = Server.MapPath("system/application/views/contents" & replace(CURRENTPATH, "\", "/"))
	
	content_load = content_file(file_path)

end function




' ---------------------------------------------------------------
'  Load External Content
' ---------------------------------------------------------------
' 
'  @string

function content_file(file_path)
	
	dim file_system, file
	
	dim local_str
	
	Set file_system	= Server.CreateObject("Scripting.FileSystemObject")
	
	Set file		= file_system.OpenTextFile(file_path, 1)
		
	do while file.AtEndOfStream = false
		
		local_str = local_str & content_parse_asp(file.ReadLine) & vbcrlf
		
	loop
	
	file.Close
	
	Set file = Nothing
	
	Set file_system = Nothing
	
	content_file = local_str

end function




' ---------------------------------------------------------------
'  Parse Simple ASP code loaded in the content
' ---------------------------------------------------------------
' 
'  @string

function content_parse_asp(inLine)
	
	dim local_str
	dim l_ia, l_ib, l_sLine, l_sLine2, l_sToken, newValue
	
	l_ia = instr(inLine, "<%")
	
	if  instr(inLine, "<%") > 0 then
		
		l_sLine 	= left(inLine, l_ia - 1)
		l_ib 		= instr(inLine, "%"&">")
		l_sLine2 	= mid(inLine, l_ib + 2)
		l_sToken 	= trim(mid(inLine, l_ia + 2, (l_ib - l_ia - 2)))
		
		local_str = Replace(eval(l_sToken),"<%","")
		
	else
	
		local_str = inLine
	
	end if
		
	content_parse_asp = local_str
	
end function

%>