<%
	'Dependencies: OutputHelper.asp

	Dim error
	Set error = New ErrorHelper

	Class ErrorHelper
		Public Function Handle(errorCode)
			If (Request.ServerVariables("HTTP_X-Requested-With") = "XMLHttpRequest") Then
				Select Case errorCode
					Case CSRF_ERROR_CODE
						Response.Redirect("/Views/Errors/InvalidRequestError.asp")
					Case Else
						Response.Redirect("/Views/Errors/UnknownError.asp")
				End Select
			Else
				' TODO: generate error output for ajax requests
				output.WriteLine("An error occured.")
			End If
		End Function

		Public Function ThrowError
			'Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(header) + "] header was unexpected"
		End Function

		Public Function getErrors
			'clear buffer'
			'write error message
			'response flush'
			GetErrors = ""
		End Function

		Public Function GetCustomErrors
			'Custom Error Handling
			If (Session("error") <> "") Then
				'todo: display a pupup or inline error message
				GetCustomErrors = ""

				'clear the error variable
				Session("error") = ""
			End If
		End Function

		Public Function GetAspErrors
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

			GetAspErrors = ""

			If (Err.Number <> 0) Then
			     Err.Clear
			End If
		End Function

		Public Function GetStingerErrors
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

		    GetStingerErrors = returnValue
		End Function
	End Class
%>