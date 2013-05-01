<%
	Dim MessageBox
	Set MessageBox = New cPageMessageBox
	
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::PAGE MESSAGEBOX
	
	Class cPageMessageBox
		
		Public Function Alert(strMessage)
			Dim sMsg
			sMsg = Replace(strMessage,"'","\'")
			sMsg = Replace(sMsg,vbCr,"\n")
			sMsg = "<script>alert('"+ sMsg +"');</script>"				
			Page.RegisterClientScript "showalert",sMsg		
		End Function
		
		Public Function Confirm(strMessage)
			Dim mScript
			mScript = "<script>" + _
				"if (confirm('"+ cSingleLineText(strMessage) +"')) " + _
				"	clasp.form.doPostBack('OnMessageConfirm','Page','OK');" + _
				"</script>"
			Page.RegisterClientScript "showconfirm",mScript
		End Function

		Public Function Prompt(strMessage,strDefault)
			Dim mScript
			mScript = "<script> var vl;" + _
				"if (vl=prompt('"+ cSingleLineText(strMessage) +"','"+ cSingleLineText(strDefault) +"')) " + _
				"	clasp.form.doPostBack('OnMessagePrompt','Page',vl);" + _
				"</script>"
			Page.RegisterClientScript "showprompt",mScript
		
		End Function

		Public Function Show(strTitle,strMessage,x,y,w,h)
			Dim sMsg

			Set sMsg = New StringBuilder
			sMsg.Append "<div id='__msg' class='thedummyclass' style='border:groove 2px #0033cc; padding:3px; font-size:10pt;position: absolute; left: " & x & ";width:" & w & ";top:" & y & ";height:" & h & ";font-family: tahoma;background-color:beige;' onmouseover='___msg=this'>"
			sMsg.Append "<table bgcolor='#E0E0E0' width='100%' border='0' ID='Table1'><tr><td style='font-weight:bold;font-family:tahoma;color:blue'>" & strTitle & "</td><td align=right><a href=""javascript:;"" onclick="" ___msg.style.visibility ='hidden';"">close</a></td></tr></table><hr size='1'>"
			sMsg.Append strMessage
			sMsg.Append "</div>"								
			Page.RegisterClientScript "showmessage",sMsg.ToString()
		End Function

		Private Function cSingleLineText(t)
			t = Replace(t,"'","\'")
			t = Replace(t,vbCr,"\n")		
			cSingleLineText = t
		End Function
	End Class
%>
