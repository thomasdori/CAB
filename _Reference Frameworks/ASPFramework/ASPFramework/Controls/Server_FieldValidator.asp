<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Validator
	Page.RegisterClientStartupScript "ServerValidator","<script language=javascript>clasp.loadLibrary('validation');</script>"
	Public Function New_ServerRequiredFieldValidator(name,control,msg,init,group) 
		Set New_ServerRequiredFieldValidator = New ServerValidator
		With New_ServerRequiredFieldValidator
			.Control.Name = name
			.InitialValue = init
			.ValidatorType = 1
			.Text = msg			
			.ValidationGroup = group
			Set .ControlToValidate = control			
		End With
	End Function

	Public Function New_ServerDateFieldValidator(name,control,group) 
		Set New_ServerDateFieldValidator = New ServerValidator
		With New_ServerDateFieldValidator
			.Control.Name = name
			.ValidationGroup = group
			.ValidatorType = 2
			Set .ControlToValidate = control			
		End With
	End Function

	Public Function New_ServerNumericFieldValidator(name,control,group) 
		Set New_ServerNumericFieldValidator = New ServerValidator
		With New_ServerNumericFieldValidator
			.Control.Name = name
			.ValidationGroup = group
			.ValidatorType = 3
			Set .ControlToValidate = control			
		End With
	End Function


	
	 Class ServerValidator
		Dim Control
		Dim Text
		Dim ErrorMsg
		Dim ValidatorType '1) Required, 3) Numeric, 5)Date
		Dim InitialValue
		Dim ControlToValidate
		Dim ValidationGroup
		
		Dim RenderTimesX		
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Text            = ""	
			ErrorMsg        = ""
			ValidatorType   = 1
			InitialValue    = ""
			ValidationGroup = ""
			RenderTimesX    = 0
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
		End Function
		
		Public Function WriteProperties(bag)
		End Function
	   
 	    Public Function HandleClientEvent(e)
	    End Function			
	   
	    Public Property Get IsValid()
	    End Property

		Private Function RenderValidator(DefaultMsg,DivID,ScriptText)
			Dim Style
			'position:relative;visibility:hidden;
			style = "position:relative;visibility:hidden;" & Control.Style			
			Response.Write "<DIV id='" & DivID & "' STYLE = '" & style & "' "
			If Control.CssClass<>"" Then 
				Response.Write " CLASS='" & CssClass & "' "
			End If			
			Response.Write ">" & IIF(Text="",DefaultMsg,Text) & "</DIV>"		
			
			Page.RegisterClientScript DivID ,ScriptText

		End Function
		
		Private Function RenderRequired()
			Dim DivID
			DivID = Control.ControlID & RenderTimesX
			RenderValidator Text,DivID,"<SCRIPT Language='JavaScript'>new claspReqFieldValidator('" & ControlToValidate.Control.ControlID  &  "','" & InitialValue & "','" & Text & "','" & DivID & "','" & ValidationGroup & "')</SCRIPT>"
		End Function	   	   
		
		Public Function RenderIsDate()			
			Dim DivID
			DivID = Control.ControlID & RenderTimesX
			RenderValidator "Invalid Date.",DivID,"<SCRIPT Language='JavaScript'>new claspDateFieldValidator('" & ControlToValidate.Control.ControlID  &  "','" & DivID & "','" & ValidationGroup & "')</SCRIPT>"			
		End Function

		Public Function RenderIsNumeric()
			Dim DivID
			DivID = Control.ControlID & RenderTimesX
			RenderValidator "Invalid Number.",DivID,"<SCRIPT Language='JavaScript'>new claspNumericFieldValidator('" & ControlToValidate.Control.ControlID  &  "','" & DivID & "','" & ValidationGroup & "')</SCRIPT>"
		End Function
		
	    Public Default Function Render()			
			Dim varStart	 
			
			If Control.Visible = False Then
				Exit Function
			End If
			
			varStart = Now
			Select Case ValidatorType
				Case 1: RenderRequired
				Case 2: RenderIsDate
				Case 3: RenderIsNumeric
			End Select			
			RenderTimesX = RenderTimesX + 1
		 	Page.TraceRender varStart,Now,Control.ControlID
		
		End Function

	End Class
%>