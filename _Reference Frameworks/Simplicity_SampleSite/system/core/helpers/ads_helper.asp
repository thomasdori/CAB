<%

' ---------------------------------------------------------------
'  Ads Helper
' ---------------------------------------------------------------
' 
' ads_create
' ads_grab_sizes


' ---------------------------------------------------------------
'  Create Add Unit
' ---------------------------------------------------------------
' 
'  @string

function create_ad(pType, pSize, pValues, pPosition)
	
	dim local_str, ad_type, ad_size
	
	ad_type	= Trim(pType)
	
	ad_size = grab_add_sizes(pSize)
	
	Select Case ad_type
	
	Case "1"
		
		local_str = "<a href=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/cc?z=publicrecord"">"
		local_str = local_str & "<img src=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/vc?z=publicrecord&dim=" & pValues(1) & """ width=""" & ad_size(0) & """ height=""" & ad_size(1) & """ border=""0"">"
		local_str = local_str & "</a>"
	
	Case "2"
		
		local_str = "<iframe width=""" & ad_size(0) & """ height=""" & ad_size(1) & """ noresize scrolling=No frameborder=0 marginheight=0 marginwidth=0 src=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/vh?z=publicrecord&dim=" & pValues(1) & "&pos=" & pPosition & """>"
		local_str = local_str & "<script language=JavaScript src=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/vj?z=publicrecord&dim=233367&pos=" & pPosition & "&abr=$scriptiniframe""></script>"
		local_str = local_str & "<noscript>"
		local_str = local_str & "<a href=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/cc?z=publicrecord&pos=" & pPosition & """>"
		local_str = local_str & "<img src=""http://rotator.adjuggler.com/servlet/ajrotator/" & pValues(0) & "/0/vc?z=publicrecord&dim=" & pValues(1) & "&pos=" & pPosition & "&abr=$imginiframe"" width=""" & ad_size(0) & """ height=""" & ad_size(1) & """ border=""0"">"
		local_str = local_str & "</a>"
		local_str = local_str & "</noscript>"
		local_str = local_str & "</iframe>"

		
	End Select
	
	create_ad = local_str
	
end function


' ---------------------------------------------------------------
'  Create Add Unit
' ---------------------------------------------------------------
' 
'  @array

function grab_add_sizes(pSize)
	
	dim local_arr
	
	Select Case pSize
	
		Case "widesky"
			local_arr = array("160","600")
			
		Case "leader"
			local_arr = array("728","90")
			
		Case "badge"
			local_arr = array("120","60")
			
	End Select
	
	grab_add_sizes = local_arr
	
end function
%>