<%


' ---------------------------------------------------------------
'  Create Default Functions for this Application
' ---------------------------------------------------------------
' 
' 
' 
	dim createCrumbs, show_breadcrumbs
	
	Set createCrumbs = New breadcrumbs
		show_breadcrumbs = createCrumbs.build()
	Set createCrumbs = Nothing
	
	
	
	
	
	
' ---------------------------------------------------------------
' Adding newlines for better readability
' ---------------------------------------------------------------
	
	site_content	= write
	
	site_content 	= Replace(site_content, "<br />", "<br />"&Vbcrlf)
	site_content 	= Replace(site_content, "</p>", "</p>"&Vbcrlf)
	site_content 	= Replace(site_content, "</li>", "</li>"&Vbcrlf)
	site_content 	= Replace(site_content, "</ul>", "</ul>"&Vbcrlf)
	site_content 	= Replace(site_content, "</tr>", "</tr>"&Vbcrlf)
	site_content 	= Replace(site_content, "</td>", "</td>"&Vbcrlf)
	site_content 	= Replace(site_content, "</table>", "</table>"&Vbcrlf)


' ---------------------------------------------------------------
'  Build Template
' ---------------------------------------------------------------
	
	dim g_oPageGen
	
	set g_oPageGen = New parseTMPL
		
		g_oPageGen.TemplateFile = "/system/application/views/templates/" & default_template
		g_oPageGen.AddToken "meta_title"			,default_site_name
		g_oPageGen.AddToken "site_url"				,config_site_url
		g_oPageGen.AddToken "page_title" 			,page_title
		g_oPageGen.AddToken "main_content" 			,site_content
		g_oPageGen.AddToken "breadcrumbs"			,show_breadcrumbs
		g_oPageGen.AddToken "show_flash_messages"	,show_flash_messages
		g_oPageGen.AddToken "page_rendered"			,formater(now(), "general date")
		g_oPageGen.AddToken "page_execute"			,page_execute(STARTTIME)
		g_oPageGen.GenerateHTML
	set g_oPageGen = nothing
	
	Response.write show_site_info("hide")
	
%>