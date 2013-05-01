<%

' ---------------------------------------------------------------
'  Post Helpers
' ---------------------------------------------------------------
' 
' post
' post_safe
' do_post
' post_show_all

' ---------------------------------------------------------------
'  post
' ---------------------------------------------------------------
' 
' @var = string

function post(pName,pDefault)
	
	post = do_post(pName,pDefault, "clean")
	
end function





' ---------------------------------------------------------------
'  post
' ---------------------------------------------------------------
' 
' @var = string

function post_safe(pName,pDefault)
	
	post = do_post(pName,pDefault, "safe")
	
end function





' ---------------------------------------------------------------
'  Do Post Function
' ---------------------------------------------------------------
' 
' @var = string

function do_post(pName,pDefault,pClean)
	
	if NOT does_exists(Request(pName)) then
		
		if pClean = "clean" then
			
			do_post = kill_illegal( pDefault )
		
		else
			do_post = pDefault
		
		end if
		
	else
		
		if pClean = "clean" then
			
			do_post = kill_illegal( Request(pName) )
		
		else
			do_post = Request(pName)
		
		end if
	
	end if
	
end function


' ---------------------------------------------------------------
'  Show All Post values
' ---------------------------------------------------------------
' 
' @var = string

function post_show_all
	
	
	If Request.Form <> "" Then
		
		Dim post_val(), createTable
    	
		ReDim post_val(Request.Form.Count - 1,1)
		
		i = 0
		
		for each Item In Request.Form
    		
			post_val(i,0) =  Item
			
			post_val(i,1) = Request.Form(Item)
			
			i = i + 1
			
		next
		
		Set createTable = New htmlTable
		
			createTable.headers 	= array("Form Name","Form Value")
			
			createTable.altRows 	= true
			
			createTable.array_data 	= post_val
			
			post_show_all = createTable.build()
			
		Set createTable = Nothing
	
	end if
	
end function

%>