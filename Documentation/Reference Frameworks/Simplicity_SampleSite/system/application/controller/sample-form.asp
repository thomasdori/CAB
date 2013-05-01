<%


' ---------------------------------------------------------------
' Content for Controller Page
' ---------------------------------------------------------------
' 
' @write = string
	
	page_title = "Sample Form"
	
	write = "<p>This is a sample form and how to use the form validation class.  Not all validation features are being used here.  Please see the user guide to see all the features</p>"
	
	Set ValForm = New ValidateRequest
	
	' Set name element
	ValForm.ElementName 	= "name"
	ValForm.ElementType 	= "text" 
	ValForm.Val_Required 	= TRUE
	'ValForm.Msg_Required 	= "Please enter your name!"
	ValForm.Val_Min = 3
	ValForm.Msg_Min = "Your name must be greater than 3 letters!"
	ValForm.Val_Max = 10
	ValForm.Msg_Max = "We only accept people with names under 10 letters ;-)"
	valForm.bool_Validate
	
	' Set email element
	ValForm.ElementName 	= "email"
	ValForm.ElementType 	= "text" 
	ValForm.Val_Required 	= TRUE
	ValForm.Msg_Required 	= "Please enter your email address!"
	ValForm.Val_Email 		= TRUE
	ValForm.Msg_Email 		= "Please enter a valid email address!"	
	emailCool = valForm.bool_Validate
	
	' Set telephone element
	ValForm.ElementName 	= "landline"
	ValForm.ElementType 	= "text" 
	ValForm.Val_Required 	= TRUE
	ValForm.Alt_Required	= "mobile,office"
	ValForm.Msg_Required 	= "Please enter at least 1 telephone number!"
	valForm.bool_Validate
	
	' Set notes element
	ValForm.ElementName 	= "notes"
	ValForm.ElementType 	= "text" 
	ValForm.Val_Required 	= TRUE
	ValForm.Msg_Required 	= "Please enter some notes!"
	valForm.bool_Validate
	
	' Set number element
	ValForm.ElementName 	= "number"
	ValForm.ElementType 	= "number" 
	ValForm.Val_Required 	= TRUE
	ValForm.Msg_Required 	= "Please enter numbers only!"
	ValForm.Val_Min = 3
	ValForm.Msg_Min = "Please enter a value greater than 3!"
	ValForm.Val_Max = 10
	ValForm.Msg_Max = "Please enter a number less than 10"
	valForm.bool_Validate
	
	' Set date element
	ValForm.ElementName 	= "date"
	ValForm.ElementType 	= "text" 
	ValForm.Val_Date 		= TRUE
	ValForm.Msg_Date 		= "Valid dates only!"
	valForm.bool_Validate
	
	
	error_count = valForm.int_error_count
	' dict_get_errors returns a dictionary object with name value/pairs of the 'element name' and 'error messages '
	Set error_messages = valForm.dict_get_errors
	' Clean Up
	Set ValForm = Nothing
	
	If Request.Form.Count > 1 Then
		If error_count = 0 Then
			' Do success code
			err_msg = show_static_messages("info","Form Validated Successfully")
		Else
			' The validation has failed
			err_msg = show_form_errors(error_count,error_messages)
			
		End If
	End If
	
	
	
	write = write & form_open("")
	
	write = write & err_msg
	
	write = write & "<fieldset>"
	
	write = write & write_form_div("Name","name",true,form_input("name","","",error_messages.item("name")))
	
	write = write & write_form_div("Email","email",true,form_input("email","","",error_messages.item("email")))
	
	write = write & write_form_div("Tel - Land Line","landline",true,form_input("landline","","",error_messages.item("landline")))
	
	write = write & write_form_div("Tel - Office","office",true,form_input("office","","",error_messages.item("office")))
	
	write = write & write_form_div("Tel - mobile","mobile",true,form_input("mobile","","",error_messages.item("mobile")))
	
	write = write & write_form_div("Notes","notes",true,form_input("notes","","",error_messages.item("notes")))
	
	write = write & write_form_div("Number","number",true,form_input("number","","",error_messages.item("number")))
	
	write = write & write_form_div(" ","submit",false,form_submit("submit","submit") & nb(3) & anchor("/sample-form","Cancel",""))
	
	write = write & "</fieldset>"
	
	write = write & form_close("<br />")
	
    