<%

' ---------------------------------------------------------------
'  Profiler Helper
' ---------------------------------------------------------------
' 
' show_profiler

' ---------------------------------------------------------------
'  Create Add Unit
' ---------------------------------------------------------------
' 
'  @string

function show_profiler
	
	dim local_str, count, Item, createTable, x, y
	
	local_str = "<br clear=""all"">" & Vbcrlf
	
	local_str = local_str & "<style>" & Vbcrlf
	local_str = local_str & "#profiler_table { font-size: 12px; font-family: arial, helvetica, sans-serif; width: 100%; margin: 2px;}" & Vbcrlf
	local_str = local_str & "#profiler_table td { font-size: 12px; font-family: arial, helvetica, sans-serif; border-bottom: 3px solid #efefef; border-left: 3px solid #efefef; width: 50%; padding:4px 6px;color:#666;background-color:#ddd; }" & Vbcrlf
	local_str = local_str & "#profiler_table th { font-size: 12px; font-family: arial, helvetica, sans-serif; border-bottom: 3px solid #efefef; width: 50%; padding:4px 6px;color:#666;background-color:#ddd; }" & Vbcrlf
	local_str = local_str & "</style>" & Vbcrlf
	
	local_str = local_str & "<div style='color: #666; background-color:#fff;padding:20px; text-align: left'>" & Vbcrlf
	
	local_str = local_str & "<strong>Application Profiler</strong> - <em>Loading Time: " & page_execute(STARTTIME) & "</em>"  & Vbcrlf
	
	
	
	'--------------------------------------------------------------
	'	Display Form Post Table
	'--------------------------------------------------------------
	
	local_str = local_str & "<fieldset style='border:1px solid #cd6e00;padding:10px 20px;margin:20px 0;background-color:#f7f7f7'>" & Vbcrlf
	
	local_str = local_str & "<legend style='padding:4px;color:#cd6e00;border:1px solid #cd6e00;background-color:#fff'><strong>&nbsp;&nbsp;Post Data&nbsp;&nbsp;</strong></legend>" & Vbcrlf
	
	If Request.Form <> "" Then
	
		Dim value4() 
    
		ReDim value4(Request.Form.Count - 1,1)
	
		i = 0
	
		for each Item In Request.Form
    		
			value4(i,0) =  Item
			
			value4(i,1) = Request.Form(Item)
			
			i = i + 1
			
		next
		
		Set createTable = New htmlTable
			
			createTable.cellpadding	= 4
			
			createTable.table_style	= "profiler_table"
			
			createTable.array_data 	= value4
			
			local_str = local_str & createTable.build()
		
		Set createTable = nothing
		
	else
	
		local_str = local_str & "<div style='color:#666;font-weight:normal;padding:4px 0 4px 0'>No Post data exists</div>" & Vbcrlf
	
	end if
	
	local_str = local_str & "</fieldset>" & Vbcrlf
	
	
	
	
	'--------------------------------------------------------------
	'	Display Cookies.
	'--------------------------------------------------------------
	
	local_str = local_str & "<fieldset style='border:1px solid #009900;padding:10px 20px;margin:20px 0;background-color:#f7f7f7'>" & Vbcrlf
	
	local_str = local_str & "<legend style='padding:4px;color:#009900;border:1px solid #009900;background-color:#fff'><strong>&nbsp;&nbsp;Cookie Data&nbsp;&nbsp;</strong></legend>" & Vbcrlf
	
	If Request.Cookies <> "" Then
		
		Dim value1() 
    	
		ReDim value1(Request.Cookies.Count - 1,1)
		
		i = 0
		
		for each x in Request.Cookies
			
			if Request.Cookies(x).HasKeys then
				
				for each y in Request.Cookies(x)
					
					value1(i,0) =  x & ":" & y
					value1(i,1) = Replace(Request.Cookies(x)(y),"%","% ")
					
				next
				
			else
				
				value1(i,0) =  x
				value1(i,1) =  Replace(Request.Cookies(x),"%","% ")
				
			end if
			
			i = i + 1
			
		next
			
			Set createTable = New htmlTable
				
				createTable.cellpadding	= 4
				
				createTable.table_style	= "profiler_table"
				
				createTable.array_data 	= value1
				
				local_str = local_str & createTable.build()
				
			Set createTable = nothing
	else
		
		local_str = local_str & "<div style='color:#666;font-weight:normal;padding:4px 0 4px 0'>No Cookie data exists</div>" & Vbcrlf
	
	end if
	
	local_str = local_str & "</fieldset>" & Vbcrlf
	
	
	
	
	'--------------------------------------------------------------
	'	Display Session.
	'--------------------------------------------------------------
	
	local_str = local_str & "<fieldset style='border:1px solid #0000FF;padding:10px 20px;margin:20px 0;background-color:#f7f7f7'>" & Vbcrlf
	
	local_str = local_str & "<legend style='padding:4px;color:#0000FF;border:1px solid #0000FF;background-color:#fff'><strong>&nbsp;&nbsp;Session Data&nbsp;&nbsp;</strong></legend>" & Vbcrlf
	
	If Session.Contents.Count <> 0 Then
	
		count=Session.Contents.Count - 1
	
		Dim value3() 
    	
		ReDim value3(count,1)
		
		i = 0
		
		for each Item in Session.Contents
			value3(i,0) =  Item
			
			value3(i,1) = Session.Contents(Item)
			
			i = i + 1
			
		next
		
		Set createTable = New htmlTable
			
			createTable.cellpadding	= 4
			
			createTable.table_style	= "profiler_table"
			
			createTable.array_data 	= value3
			
			local_str = local_str & createTable.build()
		
		Set createTable = Nothing
		
	else
	
		local_str = local_str & "<div style='color:#666;font-weight:normal;padding:4px 0 4px 0'>No Session data exists</div>" & Vbcrlf
		
	end if
	
	local_str = local_str & "</fieldset>" & Vbcrlf
	
	
	
	'--------------------------------------------------------------
	'	Display Server Varibles.
	'--------------------------------------------------------------
	
	local_str = local_str & "<fieldset style='border:1px solid #ff0000;padding:10px 20px;margin:20px 0;background-color:#f7f7f7'>" & Vbcrlf
	
	local_str = local_str & "<legend style='padding:4px;color:#ff0000;border:1px solid #ff0000;background-color:#fff'><strong>&nbsp;&nbsp;Server Varibles Data&nbsp;&nbsp;</strong></legend>" & Vbcrlf
	
	count = Request.ServerVariables.Count - 1
	
	Dim value2() 
    
	ReDim value2(count,1)
	
	i = 0
	
	For Each Item in Request.ServerVariables
   		
		value2(i,0) =  Item
		
		value2(i,1) = Replace(Request.ServerVariables(Item),"=","= ")
		
		i = i + 1
	
	Next
	
	Set createTable = New htmlTable
		
		createTable.table_style	= "profiler_table"
		
		createTable.array_data 	= value2
		
		local_str = local_str & createTable.build()
	
	Set createTable = nothing
	
	
	local_str = local_str & "</fieldset>" & Vbcrlf
	
	local_str = local_str & "</div>" & Vbcrlf
	
	show_profiler = local_str
	
end function
%>