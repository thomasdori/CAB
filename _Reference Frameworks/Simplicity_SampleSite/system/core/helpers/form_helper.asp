<%

' ---------------------------------------------------------------
'  Form Helpers
' ---------------------------------------------------------------
' 
' form_open
' form_close
' form_submit
' form_hidden
' form_input
' form_open_multipart
' form_dropdown
' form_multiple_dropdown
' form_password
' form_checkbox
' form_radio
' write_form_div
' write_form_td

' ---------------------------------------------------------------
'  Open Form
' ---------------------------------------------------------------
' 
' @var = string
function form_open(pURL)
	
	dim action
	
	if does_exists(pURL) then
	
		action = " action=""" & pURL & PAGEEXT & """"
	
	end if
	
	form_open = "<form method=""post"""&action&">"

end function




' ---------------------------------------------------------------
'  Open Close
' ---------------------------------------------------------------
' 
' @var = string
function form_close(pEnd)
	
	form_close = "</form>" & pEnd

end function




' ---------------------------------------------------------------
'  Form Submit
' ---------------------------------------------------------------
' 
' @var = string
function form_submit(pName,pValue)
	
	form_submit = "<input class=""btn"" type=""submit"" name="""&pName&""" value="""&pValue&""" />"

end function






' ---------------------------------------------------------------
'  Form Hidden
' ---------------------------------------------------------------
' 
' @var = string
function form_hidden(pName,pValue)
	
	form_hidden = "<input type=""hidden"" name="""&pName&""" value="""&pValue&""" />"

end function






' ---------------------------------------------------------------
'  Form Input
' ---------------------------------------------------------------
' 
' @var = string
function form_input(pName,pValue,pAtts,pValid)
	
	dim valid_class, valid_img
	
	if isArray(pAtts) then
		
		pAtts = add_attributes(pAtts)
	
	end if
	
	pValue = form_text_value(pName,pValue)
	
	valid_class = create_form_error(pValid,"class")
	
	valid_img = create_form_error(pValid,"img")
	
	form_input = "<input type=""text"" id="""&pName&""" name="""&pName&""" value="""&pValue&""" "&pAtts&" "&valid_class&" />" & valid_img

end function






' ---------------------------------------------------------------
'  Form Grab Values Text
' ---------------------------------------------------------------
' 
' @var = string
function form_text_value(pName,pValue)
	
	dim plocal
	
	if does_exists(Request(pName)) then
	
		plocal = Request(pName)
	
	else
		plocal = pValue
	
	end if
	
	if does_exists(plocal) then
	
		plocal = Replace(plocal,"""","&quot;")
	
	end if
	
	form_text_value = plocal

end function






' ---------------------------------------------------------------
'  Form Input
' ---------------------------------------------------------------
' 
' @var = string
function form_open_multipart(pURL)
	
	dim action
	
	if does_exists(pURL) then
	
		action = " action=""" & pURL & PAGEEXT & """"
	
	end if
	
	form_open = "<form method=""post"""&action&"  enctype=""multipart/form-data"" />"
	
end function






' ---------------------------------------------------------------
'  Password Input
' ---------------------------------------------------------------
' 
' @var = string
function form_password(pName,pValue,pAtts,pValid)
	
	dim valid_class, valid_img
	
	if isArray(pAtts) then
	
		pAtts = add_attributes(pAtts)
	
	end if
	
	pValue = form_text_value(pName,pValue)
	
	valid_class = create_form_error(pValid,"class")
	
	valid_img = create_form_error(pValid,"img")
	
	form_password = "<input type=""password"" id="""&pName&""" name="""&pName&""" value="""&pValue&""" "&pAtts&" "&valid_class&" />" & valid_img

end function






' ---------------------------------------------------------------
'  Dropdown Box
' ---------------------------------------------------------------
' 
' @var = string
function form_dropdown(pName, pValues, pValue, pAtts, pValid)
	
	dim plocal, valid_class, valid_img
	
	if isArray(pAtts) then
	
		pAtts = add_attributes(pAtts)
	
	end if
	
	valid_class = create_form_error(pValid,"class")
	
	valid_img = create_form_error(pValid,"img")
	
	plocal = "<select id="""&pName&""" name="""&pName&""""&pAtts&" " & valid_class & ">" & Vbcrlf
	
	for i = 0 to ubound(pValues)
		
		if ucase(pValues(i)) = ucase(form_text_value(pName,pValue)) then
		
			plocal = plocal & "<option value=""" & pValues(i) & """ selected>" & pValues(i+1) & "</option>" & Vbcrlf
		
		else
			
			plocal = plocal & "<option value=""" & pValues(i) & """>" & pValues(i+1) & "</option>" & Vbcrlf
		
		end if
		
		i = i + 1
	
	next
	
	plocal = plocal & "</select>" & valid_img & Vbcrlf
	
	form_dropdown = plocal

end function






' ---------------------------------------------------------------
'  Dropdown Box with Multiple Values
' ---------------------------------------------------------------
' 
' @var = string

function form_multiple_dropdown(pName,pValues,pValue,pAtts)
	
	dim plocal, arrRequest, arrValue
	
	if isArray(pAtts) then
	
		pAtts = add_attributes(pAtts)
	
	end if
	
	if NOT does_exists(request(pName)) then
		arrRequest 	= Split(pValue, ",")
	else
		arrRequest 	= Split(request(pName), ",")
	end if
	
	
	
	plocal = "<select id="""&pName&""" name="""&pName&""""&pAtts&" multiple=""multiple"">" & Vbcrlf
	
	for i = 0 to ubound(pValues)
		
		if does_exists(array_check(arrRequest, pValues(i))) then
		
			plocal = plocal & "<option value=""" & pValues(i) & """ selected=""selected"">" & pValues(i+1) & "</option>" & Vbcrlf
		
		else
			
			if NOT does_exists(request(pName)) and i = 0 and NOT does_exists(pValue)then
				
				plocal = plocal & "<option value=""" & pValues(i) & """ selected=""selected"">" & pValues(i+1) & "</option>" & Vbcrlf
			
			else
				
				plocal = plocal & "<option value=""" & pValues(i) & """>" & pValues(i+1) & "</option>" & Vbcrlf
				
			end if
			
		end if
		
		i = i + 1
	
	next
	
	plocal = plocal & "</select>" & Vbcrlf
	
	form_multiple_dropdown = plocal

end function





' ---------------------------------------------------------------
'  Grab Value from an Array
' ---------------------------------------------------------------
' 
' @var = string
function form_array_value(pName,pValue,pType)
	
	dim plocal
	
	if does_exists(Request(pName)) then
	
		plocal = Request(pName)
	
	else
	
		plocal = pValue
	
	end if
	
	plocal = Replace(plocal,"""","&quot;")
	
	form_text_value = plocal
	
end function






' ---------------------------------------------------------------
'  Textarea Box
' ---------------------------------------------------------------
' 
' @var = string
function form_textarea(pName,pValue,pAtts)
	
	if isArray(pAtts) then
		
		pAtts = add_attributes(pAtts)
	
	end if
	
	pValue = form_text_value(pName,pValue)
	
	form_textarea = "<textarea id="""&pName&""" name="""&pName&""" "&pAtts&">"&pValue&"</textarea>"
	
end function






' ---------------------------------------------------------------
'  Check Box
' ---------------------------------------------------------------
' 
' @var = string
function form_checkbox(pName,pValue,pSelected,pAtts)
	
	dim isChecked
	
	if isArray(pAtts) then
		
		pAtts = add_attributes(pAtts)
	
	end if
	
	if does_exists(request(pName)) then
		
		if pValue = form_text_value(pName,pValue) then
			
			isChecked = "checked=""checked"" "
		
		end if
	
	else
	
		if pValue = pSelected then
			
			isChecked = "checked=""checked"" "
		
		end if
		
	end if
	
	form_checkbox = "<input type=""checkbox"" id="""&pName&""" name="""&pName&""" "&pAtts&" value="""&pValue&""" "&isChecked&" />"
	
end function






' ---------------------------------------------------------------
'  Radio Button
' ---------------------------------------------------------------
' 
' @var = string
function form_radio(pName,pValue,pSelected,pAtts)
	
	dim isChecked
	
	if isArray(pAtts) then
		
		pAtts = add_attributes(pAtts)
	
	end if
	
	if does_exists(request(pName)) then
	
		if pValue = form_text_value(pName,pValue) then
		
			isChecked = "checked=""checked"" "
		
		end if
	else
	
	 if pSelected then
	 
	 	isChecked = "checked=""checked"" "
	 
	 end if
	
	end if
	
	form_radio = "<input type=""radio"" id="""&pName&""" name="""&pName&""" "&pAtts&" value="""&pValue&""" "&isChecked&" />"
	
end function






' ---------------------------------------------------------------
'  Write Form Div
' ---------------------------------------------------------------
' 
' @var = string
function write_form_div(pLabel,pName,pReg,pForm)
	
	dim local_html
	
	local_html = "<div class=""form_text"">" & Vbcrlf
	
	local_html = local_html & "<label for="""&pName&""">" & pLabel
	
	if pReg then
		
		local_html = local_html & "<span class=""required"">*</span>" & Vbcrlf
	
	end if
	
	local_html = local_html & "</label>" & Vbcrlf
	
	local_html = local_html & pForm
	
	local_html = local_html & "</div>" & Vbcrlf
	
	write_form_div = local_html

end function







' ---------------------------------------------------------------
'  Write Form Div
' ---------------------------------------------------------------
' 
' @var = string
function write_form_td(pLabel,pName,pReg,pForm)
	
	dim local_html
	
	local_html = "<tr>" & Vbcrlf
	
	local_html = local_html & "<th scope=""row"" class=""forms"">" & Vbcrlf
	
	local_html = local_html & pLabel
	
	if pReg then
		
		local_html = local_html & "<span class=""required"">*</span>" & Vbcrlf
	
	end if
	
	local_html = local_html & "</th>" & Vbcrlf
	
	local_html = local_html & "<td class=""forms"">"  & pForm  & "</td>" & Vbcrlf
	
	local_html = local_html & "</tr>" & Vbcrlf
	
	write_form_td = local_html

end function







' ---------------------------------------------------------------
'  Create the form error Style
' ---------------------------------------------------------------
' 
' @var = string

function create_form_error(pValue,pType)
	
	if does_exists(pValue) then
		
		if pType = "class" then
		
			create_form_error = " class=""invalid"""
		
		else
		
			create_form_error = "<img src='/public/imgs/icons/stop.png' alt='' width='16' height='16' border='0'  style='padding: 0 5px' />"
		
		end if
	
	end if
	
	
end function


%>