<%

' ---------------------------------------------------------------
'  XML Helpers
' ---------------------------------------------------------------
' 
' load XML




' ---------------------------------------------------------------
'  Load XML
' ---------------------------------------------------------------
' 
' @return = Collection

function load_xml(pUrl,pNode)
	
	dim xmlhttp, source
    
	set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	
    xmlhttp.Open "GET", pUrl, false
    
	xmlhttp.Send
	
	set source = Server.CreateObject("Microsoft.XMLDOM")
	
	source.async = false 
	
	source.loadxml(xmlhttp.ResponseText)
	
    set MemberNodes = source.documentElement.getElementsByTagName(pNode)
	
	set load_xml = MemberNodes
	
end function

%>