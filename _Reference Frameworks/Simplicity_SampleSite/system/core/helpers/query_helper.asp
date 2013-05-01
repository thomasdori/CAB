<%

' ---------------------------------------------------------------
'  Query Helpers
' ---------------------------------------------------------------
' 
' validate_sql
' count_all
' count_where
' count_sql
' select_all
' select_where
' update
' insert
' delete
' add_where_sql
' add_insert_sql
' debug_sql
' get_single_value
' create_sql_array


function validate_sql(sql_stmt)
	
	on error resume next
	
	dim local_conn, local_rs
	
	set local_conn = db_open(config_db)
	
	set local_rs = local_conn.Execute("SET PARSEONLY ON; " & sql_stmt & "")
	
	if err.Number = 0 then
		
		validate_sql = true
		
	else
		
		validate_sql = false
		
	end if
	
	call db_close(local_conn)
	
	set local_conn = db_open("15")
	
	set local_rs = local_conn.Execute("SET PARSEONLY OFF;")
	
	call db_close(local_conn)
	
end function



' ---------------------------------------------------------------
'  Count All
' ---------------------------------------------------------------
' 

function count_all(var)
	
	count_all = "select count(*) as count_all from [" & var & "] "
	
end function



' ---------------------------------------------------------------
'  Count Table
' ---------------------------------------------------------------
' 

function count_table(table)
	
	dim sql_conn, local_sql, local_rs
	
	Set sql_conn = db_open(config_db)
	
	Set local_rs = sql_conn.Execute( count_all(table) )
	
	if NOT local_rs.EOF then
		
		count_table = local_rs("count_all")
		
	end if
	
	call db_close(local_rs)
	
	call db_close(sql_conn)	
	
	
	
end function

' ---------------------------------------------------------------
'  Count Where
' ---------------------------------------------------------------
' 

function count_where(var, attributes)
	
	count_where = "select count(*) as count_all from [" & var & "] " & add_where_sql(attributes)
	
end function



' ---------------------------------------------------------------
'  Count FROM SQL
' ---------------------------------------------------------------
' 

function count_sql(sql)
	
	dim sql_conn, local_sql, local_rs
	
	local_sql = "select count(*) as record_count from (" & sql & ") as tbl_1"
	
	Set sql_conn = db_open(config_db)
	
	'debug_sql(local_sql)
	
	Set local_rs = sql_conn.Execute(local_sql)
	
	if NOT local_rs.EOF then
		
		count_sql = local_rs("record_count")
		
	end if
	
	call db_close(local_rs)
	
	call db_close(sql_conn)	
	
end function





' ---------------------------------------------------------------
'  Select All
' ---------------------------------------------------------------
' 

function select_all(var)
	
	select_all = "select * from " & var
	
end function





' ---------------------------------------------------------------
'  Select Where
' ---------------------------------------------------------------
' 

function select_where(var, where)
	
	select_where = "select * from " & var & " " & add_where_sql(where)
	
end function





' ---------------------------------------------------------------
'  Update
' ---------------------------------------------------------------
' 

function update(var, data, where)
	
	update = "update [" & var & "] " & add_update_sql(data) & " " & add_where_sql(where)
	
end function






' ---------------------------------------------------------------
'  Delete
' ---------------------------------------------------------------
' 

function delete(var, where)
	
	delete = "delete from [" & var & "] " & add_where_sql(where)
	
end function




' ---------------------------------------------------------------
'  Insert
' ---------------------------------------------------------------
' 

function insert(var, data)
	
	insert = "insert into  [" & var & "] " & add_insert_sql(data)
	
end function


' ---------------------------------------------------------------
'  Create Where Clause
' ---------------------------------------------------------------
' 

function add_where_sql(var)
	
	if does_exists(var) then
	
		dim attributes, atts
		
		attributes = var
		
		atts = "where " 
		
		for i = 0 to ubound(attributes)
			
			if i <> 0 then
			
				atts = atts & " and " 
			
			end if
			
			'if isnumeric(attributes(i+1)) then
			'
			'	atts = atts & attributes(i) & " = " & attributes(i+1) & " "
			'
			'else
			
				atts = atts & attributes(i) & " = '" & attributes(i+1) & "' "
			
			'end if
			
			i = i + 1
		
		next
		
		
		add_where_sql = atts
	
	else
	
		add_where_sql = ""
	
	end if
	
end function


' ---------------------------------------------------------------
'  Create Update Clause
' ---------------------------------------------------------------
' 

function add_update_sql(var)
	
	dim attributes, atts, field
	
	attributes = var
	
	atts = "set " 
	
	for i = 0 to ubound(attributes)
		
		if i <> 0 then
		
			atts = atts & ",  " 
		
		end if
		
		field = "[" & attributes(i) & "]"
		
		if NOT does_exists(attributes(i+1)) then
		
			atts = atts & field & " = NULL "
		
		elseif attributes(i+1) = "getDate()" then
		
			atts = atts & field & " = " & attributes(i+1) & " "
		
		elseif isnumeric(attributes(i+1)) then
		
			atts = atts & field & " = " & attributes(i+1) & " "
		
		elseif trim(lcase(attributes(i+1))) = "getdate()" then
			
			atts = atts & field & " = " & attributes(i+1) & " "
			
		else
			
			atts = atts & field & " = '" & replace(attributes(i+1),"'", "''") & "' "
		
		end if
		
		i = i + 1
	
	next
	
	add_update_sql =  atts
	
end function



' ---------------------------------------------------------------
'  Create insert Clause
' ---------------------------------------------------------------
' 

function add_insert_sql(var)
	
	dim attributes, atts
	
	attributes = var
	
	atts = "([" 
	
	for i = 0 to ubound(attributes)
		
		if i <> 0 then
		
			atts = atts & "],  [" 
		
		end if
		
		atts = atts & attributes(i)
		
		i = i + 1
	
	next
	
	atts = atts & "]) values ("
	
	for i = 0 to ubound(attributes)
		
		i = i + 1
		
		if i <> 1 then
		
			atts = atts & ",  " 
		
		end if
		
		if not does_exists( trim(attributes(i)) ) then
		
			atts = atts & " NULL "
		
		elseif attributes(i) ="getDate()" then
		
			atts = atts & attributes(i)
		
		elseif isnumeric(attributes(i)) then
		
			atts = atts & attributes(i)
		
		else
			
			atts = atts & " '" & replace(attributes(i),"'", "''") & "' "
		
		end if
		
		
	next
	
	atts = atts & ")"
	
	add_insert_sql =  atts
	
end function





' ---------------------------------------------------------------
'  Create insert Clause
' ---------------------------------------------------------------
' 

function debug_sql(sql)
	
	debug_sql = "<div style=""color: #666; background: #f7f7f7;line-height:16px;margin: 20px;border: 1px solid #ccc;font-family: arial, helvetica, sans-serif;  padding: 20px;font-size: 11px; "">"
	
	debug_sql = debug_sql & "<h1>Debug SQL:</h1>"
	
	debug_sql = debug_sql & "<div id=""copy_sql"">" & sql & "</div>"
	
	debug_sql = debug_sql & "</div>"
	
	Response.write debug_sql
	
end function







' ---------------------------------------------------------------
'  Get Single Value from a table
' ---------------------------------------------------------------
' 

function get_single_value(table,db,field,where)
	
	dim sql_conn, local_sql, local_rs
	
	Set sql_conn = db_open(db)
	
	local_sql = "Select " & field & " as new_value from " & table & " " & add_where_sql(where)
	
	'debug_sql(local_sql)
	
	Set local_rs = sql_conn.Execute(local_sql)
	
	if not local_rs.EOF then
		
		get_single_value = local_rs("new_value")
		
	end if
	
	call db_close(local_rs)
	
	call db_close(sql_conn)
	
end function






' ---------------------------------------------------------------
'  Create Array from SQL Statement
' ---------------------------------------------------------------
' 

function create_sql_array(sql, order, db, please_select)
	
	dim sql_conn, local_sql, local_rs, sql_array, count, i, array_count
	
	sql_array = Array()
	
	if does_exists(please_select) then
		i = 2
		count = 1
	else
		i = 0
		count = 1
	end if
	
	array_count =  (count_sql(sql,db) * 2) + count
	
	Redim sql_array( array_count )
	
	if does_exists(please_select) then
		sql_array(0) = ""
		sql_array(1) = please_select
	end if
	
	Set sql_conn = db_open(db)
	
	Set local_rs = sql_conn.Execute(sql & " ORDER BY " & order)
	
	do while NOT local_rs.EOF
		
		sql_array(i) = local_rs("VALUE")
		sql_array(i+1) = local_rs("NAME")
		
	local_rs.MoveNext
	i = i + 2
	loop
	
	call db_close(local_rs)
	
	call db_close(sql_conn)
	
	create_sql_array = sql_array
	
end function

%>