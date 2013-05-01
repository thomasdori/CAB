<%
' V 1.12
Class ValidateRequest

	Private  Form_Name
	Private  Element, Element_Type
	
	Private  Require, Require_msg, Ary_Alt_Required, Ary_Compare
	Private  Exact, Exact_msg
	Private  Matchs, Matchs_msg
	Private  Captcha, Captcha_msg
	Private  Email, Email_msg
	Private  RegularExpression, RegularExpression_msg
	Private  Min, Min_msg
	Private  Max, Max_msg
	Private  ElDate, ElDate_msg
	Private  ErorrDictionary
	Private  Error_Count
	Private  ispostback
	Private  JSValCode, doClientVal
	Private	 Collection
	Private  UsersScript
	Private  SubmitButton
	
	'Intitilize Method
	Private Sub Class_Initialize()
		Set ErorrDictionary=Server.CreateObject("Scripting.Dictionary")
		ResetDefaults()
		Error_Count = 0	
		set Collection = Request.form
	End Sub
	
	Private Sub Class_Terminate()
    	Set ErorrDictionary = Nothing
    End Sub
	
	'Properties
	Public Property Let FormName(FormNamevalue)
		Form_Name=FormNamevalue
	End Property
	
	Public Property Let ElementName(ElementNamevalue)
		Element=ElementNamevalue
	End Property
	
	Public Property Let ElementType(ElementTypevalue)
		Element_Type=ElementTypevalue
	End Property
	
	Public Property let Val_Required(booleanValue)
		Require=booleanValue
		If Require_msg = "" Then Require_msg = Element & " is required!"
	End Property
	
	Public Property let Val_Captcha(booleanValue)
		Captcha=booleanValue
		If Captcha_msg = "" Then Captcha_msg = "Please ensure the security field is filled out correctly."
	End Property
	
	Public Property let Msg_Required(msg)
		Require_msg=msg
	End Property
	
	Public Property let Alt_Required(ReqAlt)
		Ary_Alt_Required = Split(ReqAlt,",")
	End Property
	
	Public Property let Compare(ReqAlt)
		Ary_Compare = Split(ReqAlt,",")
	End Property
	
	Public Property let Val_Email(booleanValue)
		Email=booleanValue
	End Property
	
	Public Property let Msg_Email(msg)
		Email_msg=msg
	End Property
	
	Public Property let Val_Exact(ExactNumber)
		Exact=ExactNumber
		 If Exact_Msg = "" Then Exact_Msg="Please enter a exact number of "&Cstr(Exact)&" characters in length"
	End Property
	
	Public Property let Msg_Exact(msg)
		Exact_Msg=msg
	End Property
	
	
	Public Property let Val_Matchs(ElementNamevalue)
		Matchs=ElementNamevalue
		 If Matchs_Msg = "" Then Matchs_Msg="Please enter a exact number of "&Cstr(Exact)&" characters in length"
	End Property
	
	Public Property let Msg_Matchs(msg)
		Matchs_Msg=msg
	End Property
	
	Public Property let Val_Min(MinNumber)
		Min=MinNumber
		 If Max_Msg = "" Then Min_Msg="Please enter a minimum of "&Cstr(Min)&" "
	End Property
	
	Public Property let Msg_Min(msg)
		Min_Msg=msg
	End Property
	
	Public Property let Val_Max(MaxNumber)
		 Max=MaxNumber
		 If Max_Msg = "" Then Max_Msg="Please enter a maximum of "&Cstr(Max)&" "
	End Property
	
	Public Property let Msg_Max(msg)
		 Max_Msg=msg
	End Property
	
	Public Property let Val_RegExp(RExpression)
		 RegularExpression=RExpression
	End Property
	
	Public Property let Msg_RegExp(msg)
		 RegularExpression_msg=msg
	End Property
	
	Public Property let Val_Date(booleanValue)
		 ElDate=booleanValue
	End Property
	
	Public Property let Msg_Date(msg)
		 ElDate_msg=msg
	End Property
	
	Public Property set SetCollection(Object)
		set Collection = Object
	End Property		
	
	Public Property let SetCustomJavascript(JavascriptString)
		 UsersScript=JavascriptString
	End Property
	
	Public Property let SubmitOnce(element)
		 SubmitButton=element
	End Property
	
	'Public Methods
	
	Public Property get dict_get_errors()
		set dict_Get_Errors=ErorrDictionary
	End Property		
	
	Public Property get int_error_count()
		 If ispostback Then int_error_count = Error_Count
	End Property
	
	Public Function bool_Validate()
		If Collection.Count > 0 Then		
			ispostback = True
		Else
			ispostback = False
		End If
		
		
		Dim ReturnValue
		Dim Error_msg
		ReturnValue = True
		
		Error_msg = DoValidate()
		
		If ispostback Then
			If Error_msg <> ""  Then
			    Error_msg=Error_msg
				call Dict_store(Element,Error_msg)
				ReturnValue = False
			End If
		End If	
		
		ClearAll()	
		ResetDefaults()
		bool_Validate=ReturnValue
		
	End Function
	
	Public function str_client_script(funname)
		doClientVal = True
		str_client_script = GenrateJavascript(funname)
	End function
	
	'Private Methods
		
	Private Function DoValidate()
		
		Dim ElementValue
		Dim Error_msg
		If ispostback Then
			ElementValue = Collection(Element)
		End If
		
		If Captcha Then
			if ucase(replace(ElementValue," ", "")) <> ucase(session_read("captcha_session")) then
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & Captcha_msg	& "</li>"
			end if
		end if
		
		If Require Then
			
			If Element_Type = "number" Then
				If NOT IsNumeric(ElementValue) Then
					' increase error count  
					' Log msg
					Error_count=Error_count+1
					Error_msg = Error_msg & "<li>" & Require_msg	& "</li>"
				End IF
			Else
				If isArray(Ary_Alt_Required) Then
					
					Dim altReq, HasVal
					HasVal = False
					For each altReq in Ary_Alt_Required
						If ElementValue <> "" Then HasVal = True
						If Collection(altReq) <> "" Then
							HasVal = True
						End IF
					Next
					If NOT HasVal Then
						Error_count=Error_count+1
						Error_msg = Error_msg & "<li>" & Require_msg	& "</li>"
					End If
			Else
					If IsNull(ElementValue) OR ElementValue = " " OR ElementValue = "" Then
						' increase error count  
						' Log msg
						Error_count=Error_count+1
						Error_msg = Error_msg & "<li>" & Require_msg	& "</li>"
					End IF
				End If
			End If
			
		End IF
		
		if Matchs <> "" then
			
			if ElementValue <> post(Matchs, "") then
			
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & Matchs_msg	& "</li>"
			end if
		end if
		
		IF Email Then
			Dim IsEmailCorrect	
			IsEmailCorrect=isRegExp("^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$",ElementValue)
			'"^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
  			
			If Not IsEmailCorrect Then
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & Email_msg & "</li>"
			End IF		

		End If
		
		If RegularExpression <> "" Then
		
			Dim IsRegExpCorrect	
			IsRegExpCorrect=isRegExp(RegularExpression,ElementValue)
  			
			If Not IsRegExpCorrect Then
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & RegularExpression_msg & "</li>"
			End IF
			
			
		End If 
		
		If Exact <> "" then
			
			If Len(ElementValue)<>Exact Then
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & Exact_msg	& "</li>"
			End If
			
		End If
		
		
		If Min <> "" Then
			
			If Element_Type = "number" Then
						
				If IsNumeric(ElementValue) Then
					If cDbl(ElementValue) < Min Then
						Error_count=Error_count+1
						Error_msg = Error_msg & "<li>" & Min_msg	& "</li>"
					End If
				End If
				
			Else
				
				If Len(ElementValue)<Min Then
					Error_count=Error_count+1
					Error_msg = Error_msg & "<li>" & Min_msg	& "</li>"
				End If
				
			End If

		End If
		
		If Max <> "" Then
		
			If Element_Type = "number" Then
			
				If IsNumeric(ElementValue) Then
					If cDbl(ElementValue) > Max Then
						Error_count=Error_count+1
						Error_msg = Error_msg & "<li>" & Max_msg	& "</li>"
					End If
				End If
				
			Else
			
				If Len(ElementValue)>Max Then
					Error_count=Error_count+1
					Error_msg = Error_msg & "<li>" & Max_msg	& "</li>"
				End If
			
			End If
			
			
		End If
		
		If ElDate Then
			If NOT isDate(ElementValue) AND ElementValue <> "" Then
				Error_count=Error_count+1
				Error_msg = Error_msg & "<li>" & ElDate_msg	& "</li>"
			End If
			
		End If
		
		DoValidate=Error_msg
	End Function
	
	Private Sub Dict_store(Element,Error_msg)
		ErorrDictionary.Add Element,Error_msg
	End Sub
	
	Private Sub ClearAll()		
		Captcha=""
		Matchs=""
		Require=""
		Require_msg=""
		Email=""
		Email_msg=""
		RegularExpression=""
		RegularExpression_msg=""
		Min=""
		Min_msg=""
		Max=""
		Max_msg=""
		Exact=""
		Exact_msg=""
		ElDate=""
		ElDate_msg=""
		Element=""
	End Sub
	
	Private Sub ResetDefaults()
		Email_msg 				= "Please Enter a Valid Email Address!"
		RegularExpression_msg	= "Invalid Input!"
		ElDate_msg				= "Please enter a valid date!"
		Captcha					= False
		Require					= False
		Email					= False
		ElDate 					= False
		'doClientVal 			= False
	End Sub
	
	Private Function isRegExp(Pattern,SearchString)
		dim isValidE
		dim regEx
		
		isValidE = True
		set regEx = New RegExp
		
		regEx.IgnoreCase = False
		
		regEx.Pattern =Pattern 
		isValidE = regEx.Test(SearchString)
		
		isRegExp = isValidE
	End Function
		
End Class
%>