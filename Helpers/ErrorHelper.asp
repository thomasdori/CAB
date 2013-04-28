<%
Class ErrorHelper
	Sub throwError
		'Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(header) + "] header was unexpected"
	End Sub

	Sub showErrors
		'clear buffer'
		'write error message
		'response flush'
	End Sub

	Sub showCustomErrors
		'Custom Error Handling
		If (Session("error") <> "") Then
			'todo: display a pupup or inline error message

			'clear the error variable
			Session("error") = ""
		End If
	End Sub

	Sub showAspErrors
		'ASP Error Handling
		Dim oError
		Set oError = Server.GetLastError()

		'todo: lock error to database
		'output.write("ASPCode: " & oError.ASPCode & "<br />")
		'output.write("ASPDescription: " & oError.ASPDescription & "<br />")
		'output.write("Category: " & oError.Category & "<br />")
		'output.write("Column: " & oError.Column & "<br />")
		'output.write("Description: " & oError.Description & "<br />")
		'output.write("File: " & oError.File & "<br />")
		'output.write("Line: " & oError.Line & "<br />")
		'output.write("Number: " & oError.Number & "<br />")
		'output.write("Source: " & oError.Source & "<br />")

		If (Err.Number <> 0) Then
		     Err.Clear
		End If
	End Sub

	Sub showStingerErrors
	    if session("PageLoadedOnce") then
		    Dim errors
		    'Perform Server Side Validation
		    set errors = validator.validate
			  ' display any errors found to the user
			  if not errors is nothing then
			    if errors.count > 0 then
				    output.write "<P>Your request did not pass Stinger validation!</P>"
				    output.write "<TABLE BORDER=1 ALIGN=CENTER BGCOLOR=YELLOW><TR><TD>"
				    output.write validator.format(errors)
				    output.write "</TD></TR></TABLE>"
			    else
				    output.write "<P>Congratulations! Your request passed Stinger validation!</P>"
	        end if
	      end if
	    end if
	    Session("PageLoadedOnce") = true
	End Sub
End Class
%>