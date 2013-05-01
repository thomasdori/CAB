<%

' ---------------------------------------------------------------
'  Common Function
' ---------------------------------------------------------------
' 
' does_exists
' create_capctha_image_id
' show_site_info
' page_execute
' kill_illegal
' is_number
' reg_search
' reg_replace
' remove_ilegal




' ---------------------------------------------------------------
'  Double check of a Varible Exists
' ---------------------------------------------------------------
' 
'  @var = string

function does_exists(var)
	
	does_exists = True 
	
	if isarray(var) then
	
	else
	
		var = trim(var)
		
		if (isEmpty(var) or isNull(var) or var = "" or var = " ") then 
		
			does_exists = False 
		
		end if
	
	end if
	
end function




' ---------------------------------------------------------------
'  Show Value
' ---------------------------------------------------------------
' 
' @var = string
' @text = string

function show_value(var, text)
	
	if does_exists(var) then
	
		show_value = text
	
	end if
	
end function








' ---------------------------------------------------------------
'  Create Capctha Session ID for Capctah Image
' ---------------------------------------------------------------
' 
' @null

function create_capctha_image_id
	
	dim intUpperBound, intLowerBound, intRangeSize, sngRandomValue, intRandomInteger, randomchr
	
	intUpperBound = 90
	
	intLowerBound = 65
	
	For i=1 to 5
	
		Randomize()
	
		intRangeSize = intUpperBound - intLowerBound + 1
	
		sngRandomValue = intRangeSize * Rnd()
	
		sngRandomValue = sngRandomValue + intLowerBound
	
		intRandomInteger = Int(sngRandomValue)
	
		randomchr = randomchr & chr(intRandomInteger)
	
	next
	
	create_capctha_image_id = randomchr
	
end function





' ---------------------------------------------------------------
'  Show Site info at the bottom of the site
' ---------------------------------------------------------------
' 
' @var = string

function show_site_info(var)
	
	dim plocal, MachineName,  pc
	
	Set pc = CreateObject("Wscript.Network") 
    
		MachineName = left(pc.ComputerName , 1)
    
	Set pc = nothing 
	
	if var = "hide" then
		
		pLocal = pLocal & "<!--" & Vbcrlf
		
		pLocal = pLocal & "	SITE INFORMATION" & Vbcrlf
		
		pLocal = pLocal & "	URL: " & Request.ServerVariables("SERVER_NAME") & " - (" & MachineName & ")" & Vbcrlf
		
		pLocal = pLocal & "	IP: " & Request.ServerVariables("LOCAL_ADDR") & Vbcrlf
		
		pLocal = pLocal & "	Page Rendered: " &  formater(now, "General Date") & Vbcrlf
		
		pLocal = pLocal & "-->"
		
	else
		
		pLocal = pLocal & "	SITE INFORMATION" & Vbcrlf
		
		pLocal = pLocal & "	URL: " & Request.ServerVariables("SERVER_NAME") & " - (" & MachineName & ")" & Vbcrlf
		
		pLocal = pLocal & "	IP: " & Request.ServerVariables("LOCAL_ADDR") & Vbcrlf
		
		pLocal = pLocal & "	Page Rendered: " &  formater(now, "General Date") & Vbcrlf
	
	end if
	
	if SHOWPROFILER then
	
		pLocal = pLocal & show_profiler
	
	end if
	
	show_site_info = pLocal
	
end function






' ---------------------------------------------------------------
'  Show the page Execution Time
' ---------------------------------------------------------------
' 
' @starttime = TIME

function page_execute(starttime)
	
	dim endtime
    
	endtime = Timer() 
	
	page_execute = round(endtime-starttime, 3) & " seconds" 
	
end function






' ---------------------------------------------------------------
'  Check to see if a value is an Int
' ---------------------------------------------------------------
' 
' @val = string
'

function is_number(val) 
	
	if isnumeric(val) then
	
		is_number = val
	
	else
	
		is_number = 0
	
	end if
	
end function 





' ---------------------------------------------------------------
'  Check to see if a value is an Int
' ---------------------------------------------------------------
' 
' @var = string
' @regex_pattern = string

function reg_search(var, regex_pattern)
	
	dim RegEx
	
	set RegEx = new RegExp
		
		RegEx.Pattern = regex_pattern
		
		RegEx.IgnoreCase = true
		
		RegEx.Global = True
		
		reg_search = RegEx.test(var)  
	
	set RegEx = nothing

end  function




' ---------------------------------------------------------------
'  Check to see if a value is an Int
' ---------------------------------------------------------------
' 
' @var = string
' @regex_pattern = string
' @replace_with = string

function reg_replace(var, regex_pattern, replace_with)
	
	dim RegEx
	
	set RegEx = new RegExp
		
		RegEx.Pattern = regex_pattern
		
		RegEx.IgnoreCase = true
		
		RegEx.Global = True
		
		if does_exists(var) then
		
			reg_replace = RegEx.Replace(var, replace_with)
		
		end if
		
	set RegEx = nothing

end  function





' ---------------------------------------------------------------
'  Check to see if a value is an Int
' ---------------------------------------------------------------
' 
' @var = string

function remove_illegal(var)
	
	remove_illegal = reg_replace(var, "[^a-zA-Z0-9]", "")
	
end function





function repalce_position(array_value, position, replace_value, add_value)
	
	dim local_str
	for i = 0 to ubound(array_value)
		
		if position = i then
			local_str = local_str & replace_value
		end if
		
		local_str = local_str & add_value & array_value(i)
		
	next
	
	repalce_position = local_str

end function
%>