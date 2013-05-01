<%
'---  ---
' parse_cls.asp
'------  ----
' (c)2000 James Q. Stansfield (james.stansfield@iridani.com)
' This code is free for use by anyone. It is meant as a learning tool and can be passed along in any format.
class parseTMPL
	'Dimension variables
	private g_sTMPLFILE
	private g_oDict
	private g_bFILE

	private sub class_Initialize
		'Create the Scripting.Dictionary Object,
		'Set the compare mode to 1 so that it is case insensitive.
		'Also flag a boolean file so we know whether our file is there or not.
		set g_oDict = createobject("Scripting.Dictionary")
		g_oDict.CompareMode = 1
		g_bFILE = FALSE
	end sub
	
	private sub class_Terminate
		'Destroy our object.
		set g_oDict = nothing
	end sub
	
	public property let TemplateFile(inFile)
		'A file & path must be specified for the routine to work.
		g_sTMPLFILE = server.mappath(inFile)
	end property
	
	private property get TemplateFile()
		TemplateFile = g_sTMPLFILE
	end property
	
	public sub AddToken(inToken, inValue)
		'This property allows us to set our tokens.
		g_oDict.Add inToken, inValue
	end sub
	
	public sub GenerateHTML
		'This is the main, and only public method of the class.
		'This method will check whether we've specified a file or not.
		'Check for the files existance if we have specified it.
		'If the file exists we will open it and dumps it's contents into an array.
		'The array is split on VbCrLf to make it more manageable.
		if len(g_sTMPLFILE) > 0 then
			dim l_oFSO, l_oFile, l_arrFile
			set l_oFSO = createobject("Scripting.FileSystemObject")
			if l_oFSO.FileExists(TemplateFile) then
				g_bFILE = TRUE
				set l_oFile = l_oFSO.OpenTextFile(g_sTMPLFILE)
				l_arrFile = split(l_oFile.ReadAll, VbCrLf)
				l_oFile.Close
				set l_oFile = Nothing
			end if
			set l_oFSO = nothing
		else
			'Filename was never passed!
		end if
		'If we have a file, we will prase it line by line
		'by sending each line to ParseTMPL()
		if g_bFILE then
			Dim l_sLine
			for each l_sLine in l_arrFile
				ParseTMPL(l_sLine)
				ParseASP(l_sLine)
			next
		else
			'Error File Not Found!
		end if
	end sub
	
	private sub ParseTMPL(inLine)
		'This is a reasonably complex piece of code that looks for any text that is bounded by '[%' and '%]'...
		'If it finds one it takes the first part of the line up to the '[%' and displays it to the user.
		'It then passes the bounded text to getToken.
		'The it send the text that resides past the '%]' to itself(ParseTMPL), that way it can check for another
		'token on the same line.
		'If the routine doesn't find a '[%' in the line it simply displays it for the user.
		dim l_ia, l_ib, l_sLine, l_sLine2, l_sToken
		
		l_ia = instr(inLine, "[%")
		if l_ia > 0 then
			l_sLine = left(inLine, l_ia - 1)
			l_ib = instr(inLine, "%]")
			l_sLine2 = mid(inLine, l_ib + 2)
			l_sToken = trim(mid(inLine, l_ia + 2, (l_ib - l_ia - 2)))
			Response.Write (l_sLine)
			getToken(l_sToken)
			ParseTMPL(l_sLine2)
		else
			Response.Write stripASPCode(inLine) & VbCrLf
		end if
		
		
	end sub
	
	private function stripASPCode(var)
		
		stripASPCode = reg_replace(var, "<%[^>]*%"&">", "")
	
	end function
	
	
	private sub ParseASP(inLine)
		
		dim l_ia, l_ib, l_sLine, l_sLine2, l_sToken, newValue
		l_ia = instr(inLine, "<%")
		if l_ia > 0 then
			l_sLine = left(inLine, l_ia - 1)
			l_ib = instr(inLine, "%"&">")
			l_sLine2 = mid(inLine, l_ib + 2)
			l_sToken = trim(mid(inLine, l_ia + 2, (l_ib - l_ia - 2)))
			
			Response.write Replace(eval(l_sToken),"<%","")
			ParseASP(l_sLine2)
		end if
		
	end sub
	
	private sub getToken(inToken)
		'This routine checks the Dictionary for the text passed to it.
		'If it finds a key in the Dictionary it Display the value to the user.
		'If not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if g_oDict.Exists(inToken) then
			Response.Write(g_oDict.Item(inToken))
		else
			'Response.Write("<!--[%" & inToken & "%]-->")
		end if
	end sub
end class
%>