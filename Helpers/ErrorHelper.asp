<%
	Dim error
	Set error = new ErrorHelper

	Class ErrorHelper
		Function throwError
			'Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(header) + "] header was unexpected"
		End Function

		Function getErrors
			'clear buffer'
			'write error message
			'response flush'
			getErrors = ""
		End Function

		Function getCustomErrors
			'Custom Error Handling
			If (Session("error") <> "") Then
				'todo: display a pupup or inline error message
				getCustomErrors = ""

				'clear the error variable
				Session("error") = ""
			End If
		End Function

		Function getAspErrors
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

			getAspErrors = ""

			If (Err.Number <> 0) Then
			     Err.Clear
			End If
		End Function

		Function getStingerErrors
		    if session("PageLoadedOnce") then
			    Dim errors, returnValue
			    returnValue = ""
			    'Perform Server Side Validation
			    set errors = validator.validate
				  ' display any errors found to the user
				  if not errors is nothing then
				    if errors.count > 0 then
					    returnValue = returnValue & "<P>Your request did not pass Stinger validation!</P>"
					    returnValue = returnValue &  "<TABLE BORDER=1 ALIGN=CENTER BGCOLOR=YELLOW><TR><TD>"
					    returnValue = returnValue &  validator.format(errors)
					    returnValue = returnValue &  "</TD></TR></TABLE>"
				    else
					    returnValue = returnValue &  "<P>Congratulations! Your request passed Stinger validation!</P>"
		        end if
		      end if
		    end if
		    Session("PageLoadedOnce") = true

		    getStingerErrors = returnValue
		End Function
	End Class
%>