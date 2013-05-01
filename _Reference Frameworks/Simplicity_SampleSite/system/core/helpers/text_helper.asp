<%

' ---------------------------------------------------------------
'  Text Helper
' ---------------------------------------------------------------
' 
' humanize
' not_humanize
' humanize_url
' limit_text
' limit_chars
' formater
' show_greeting
' widont



' ---------------------------------------------------------------
'  humanize
' ---------------------------------------------------------------
' 
' @string
' 


function humanize(var)
	
	dim pLocal, iPos
	
	iPos = 1
	
	Do While InStr(iPos, var, " ", 1) <> 0
		
		iSpace = InStr(iPos, var, " ", 1)
		
		pLocal = pLocal & UCase(Mid(var, iPos, 1))
		
		pLocal = pLocal & LCase(Mid(var, iPos + 1, iSpace - iPos))
		
		iPos = iSpace + 1
		
	Loop
	
	pLocal = pLocal & UCase(Mid(var, iPos, 1))
	
	pLocal = pLocal & LCase(Mid(var, iPos + 1))
	
	pLocal = not_humanize(pLocal)
	
	humanize = pLocal
	
End Function




' ---------------------------------------------------------------
' Don't Humanize Text
' ---------------------------------------------------------------
' 
' @string
' 

function not_humanize(var)
	
	dim bad_names, pLocal, array_var, check_value
	
	bad_names = array("DuPage","McHenry","McKinley", "DeKalb")
	
	array_var = Split(var, " ")
	
	for i=0 to ubound(array_var)
		
		check_value = array_check(bad_names, array_var(i))
		
		if does_exists( check_value ) then
		
			pLocal = pLocal & check_value & " "
		
		else
		
			pLocal = pLocal & array_var(i) & " "
		
		end if
	
	next
	
	not_humanize = pLocal
	
end function





' ---------------------------------------------------------------
' Humanize Url Name
' ---------------------------------------------------------------
' 
' @string
' 

function humanize_url(var)
	
	var = Replace(var, "-", " ")
	humanize_url =  humanize(var)
	
End Function



' ---------------------------------------------------------------
'  Limit a Text String
' ---------------------------------------------------------------
' 
' @var = string
' @num = integer

function limit_text(var, num)
	
	Dim nPos, pLocalStr
	
	var = remove_html( var )
	
	var = Replace(var, vbCrLf, " ")
	
	If Len(var) > num And num > 4 Then
    
		nPos = InStrRev(Left(var, num - 3), " ")
	
		If nPos > 0 Then
	
			pLocalStr = Left(var, nPos) & "..."
    
	    Else
	
			pLocalStr = Left(var, num - 4) & "..."
	
	    End If
    
	Else
	
		pLocalStr = Left(var, num)
    
	End If
	
	pLocalStr = replace(pLocalStr, " ...", "...")
	
	limit_text = pLocalStr
	
end function




' ---------------------------------------------------------------
'  Limit a Text String by Chars
' ---------------------------------------------------------------
' 
' @var = string
' @num = integer

function limit_chars(var, num)
	
	dim local_str
	
	local_str = left(var, num)
	
	limit_chars = local_str
	
end function




' ---------------------------------------------------------------
'  Format Dates and Numbers
' ---------------------------------------------------------------
' 
' @var = string
' @type = string

function formater(var,var_type)
	
	dim pLocal
	
	select case lcase( var_type )
		
		case "xml"
			
			pLocal = pLocal & WeekdayName(Weekday(var),true) & " "
  			
			If day(var)< 10 Then
				pLocal = pLocal & "0" & day(var) & " "
			Else
				pLocal = pLocal & day(var) & " "
			End If
  			
			pLocal = pLocal & monthName(Month(var),true) & " "
			pLocal = pLocal & Year(var) & " "
			
			if Hour(var)< 10 then
				pLocal = pLocal & "0" & Hour(var) & ":"
			else
				pLocal = pLocal & Hour(var) & ":"
			end if
			
			if minute(var)< 10 Then
				pLocal = pLocal & "0" & minute(var) & ":"
			else
				pLocal = pLocal & minute(var) & ":"
			end If
			
			if second(var)< 10 Then
				pLocal = pLocal & "0" & second(var)
			else
				pLocal = pLocal & second(var)
			end If
			
			formater = pLocal & " EST"
		
		case "general date"
		
			formater = formatDateTime(var, 0)
		
		case "long date"
		
			formater = formatDateTime(var, 1)
		
		case "short date"
		
			formater = formatDateTime(var, 2)
		
		case "long time"
		
			formater = formatDateTime(var, 3)
		
		case "short time"
		
			formater = formatDateTime(var, 4)
			
		case "sql date"
			
			if month(var) < 10 then
				formater = "0" & month(var)
			else
				formater = month(var)
			end if
			
			if day(var) < 10 then
				formater = formater & "/0" & day(var)
			else
				formater = formater & "/" & day(var)
			end if
			
			formater = formater & "/" & year(var)
			
		case "general number"
		
			formater = formatNumber(var, 0, -1)
		
		case "currency"
		
			formater = formatcurrency(var, 2)
		
		case "fixed"
		
			formater = Replace( formatNumber(var, 2, -1), ",", "" )
		
		case "standard"
		
			formater = formatNumber(var, 2, -1)
		
		case "percent"
		
			formater = formatPercent(var, 2)
		
		case "yes/no"
		
			var = cLng(var)
		
			If var = 0 then 
		
				formater = "No" 
		
			else 
		
				formater = "Yes"
		
			end if
		
		case "true/false"
		
			var = cLng(var)
		
			If var = 0 then 
		
				formater = "False" 
		
			else 
		
				formater = "True"
		
			end if
		
		case "on/off"
		
			var = cLng(var)
		
			If var = 0 then 
		
				formater = "Off" 
		
			else 
		
				formater = "On"
		
			end if
		
		case else
		
			formater = var
	end select

end function




' ---------------------------------------------------------------
'  Show Greeting
' ---------------------------------------------------------------
' 
' @null

function show_greeting

	dim dHour, dMinute, LocalStr
	
	dHour = Hour(Now)
	
	dMinute = Minute(Now())
	
	if dMinute < 10 then
	
		dMinute = "0" & dMinute
	
	else
	
		dMinute = dMinute
	
	end if
	
	If dHour < 12 Then
	
		dHour = dHour
	
		LocalStr =  "<strong>Good Morning, </strong>"  & Date() & " " & dHour & ":" & dMinute & " AM"
	
	ElseIf dHour < 17 Then
	
		dHour = dHour - 12
	
		LocalStr = "<strong>Good Afternoon, </strong>"  & Date() & " " & dHour & ":" & dMinute & " PM"
	
	Else
	
		dHour = dHour - 12
	
		LocalStr = "<strong>Good Evening, </strong>"  & Date() & " " & dHour & ":" & dMinute & " PM"
	
	End If
	
	show_greeting = LocalStr
	
end function


' ---------------------------------------------------------------
'  widont
' ---------------------------------------------------------------
' 
' @var = string

function widont(var)
	
	dim array_var
	
	if does_exists(var) then
		
		array_var = Split(var, " ")
		
		var = Replace(var, " ", "-|-", 1, ubound(array_var) - 1 , 0)
		
		var = Replace(var, " ", "&nbsp;")
		
		widont = Replace(var, "-|-", " ")
		
	end if
	
end function


%>