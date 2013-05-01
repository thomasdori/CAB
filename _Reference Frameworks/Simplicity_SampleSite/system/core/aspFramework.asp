<%
	' ---------------------------------------------------------------
	'	Load All Helpers and Libraries
	' ---------------------------------------------------------------
%>
<!-- #include virtual = "/system/core/aspIncludes.asp" -->
<%
	' ---------------------------------------------------------------
	'	Auto Load Helpers per Application
	' ---------------------------------------------------------------
%>
<!-- #include virtual = "/system/application/helpers/autoload.asp" -->
<%
	' ---------------------------------------------------------------
	'	Build Coontroller
	' ---------------------------------------------------------------
%>
<!-- #include virtual = "/system/core/helpers/controller_helper.asp" -->
<%
	' ---------------------------------------------------------------
	'	Build Tempalte & Check for Errors
	' ---------------------------------------------------------------
%>
<!-- #include virtual = "/system/application/views/build_template.asp" -->
<!-- #include virtual = "/system/core/helpers/error_helper.asp" -->
<%
	
	' ---------------------------------------------------------------
	'	Sub Routine for Building the Controller
	' ---------------------------------------------------------------

sub SystemInclude(strFile)
	
	dim objRegexp, objFSO, strErrorStyle
	
	set objRegexp = new Regexp 
	
	objRegexp.Global = true 
	
	objRegexp.MultiLine = true 
	
	objRegexp.IgnoreCase = true 
	
	set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
	
	new_file_path = Server.MapPath("/") & "\"& strFile
	
	if objFSO.Fileexists(new_file_path) Then 
	
		Set objFile = objFSO.OpenTEXTFile(new_file_path, 1) 
		
		strSource = objFile.ReadAll 
		
		strSource = Replace(strSource, Chr(60)&Chr(37), "") 
		
		strSource = Replace(strSource, Chr(37)&Chr(62), "") 
		
		objRegexp.Pattern = "^[ \t]*(//¦\')[\s\S]*?$" 
		
		strSource = objRegexp.Replace(strSource, "") 
		
		if ERRORCHECK then
			
			On Error Resume NEXT 
			
		end if
		
		Execute strSource 
		
		if err.Number > 0 then
			
			Set objFile_err = objFSO.OpenTEXTFile(Server.MapPath("system\application\errors\ASP_error.asp"), 1) 
			
			strSource_err = objFile_err.ReadAll 
			
			strSource_err = Replace(strSource_err, Chr(60)&Chr(37), "") 
			
			strSource_err = Replace(strSource_err, Chr(37)&Chr(62), "") 
			
			objRegexp.Pattern = "^[ \t]*(//¦\')[\s\S]*?$" 
			
			strSource_err = objRegexp.Replace(strSource_err, "") 
			
			Execute strSource_err 
			
			Call send_error_email(strFile, Err.Number, Err.Source, Err.Description, Err.HelpContext)
			
		end if
		
		On error GoTo 0 
		
	else
		
		Call send_error_email(strFile, Err.Number, Err.Source, Err.Description, Err.HelpContext)
		SystemInclude("system\application\errors\404_error.asp")
		
		' Shows the IIS Asp Error
		
		'Response.write "<div "&strErrorStyle&">"
		'Response.write "<p><strong>ASP Include Error</strong></p>"
		'Response.write "<p>FATAL ERROR: Include file does not exists: <strong><i>" & strFile & "</i></strong></p>"
		'Response.write "<div>"
		
	end if 

end Sub 

%>