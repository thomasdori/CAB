<!--#Include File = "Controls\WebControl.asp"        -->
<!--#Include File = "Controls\Server_Button.asp" -->
<!--#Include File = "Controls\Server_TextBox.asp"    -->
<!--#Include File = "Controls\Server_DropDown.asp" -->
<html>
<body>
<h2>Dynamic Controls Sample</h2><hr>
<%Page.Execute%>
<%Page.OpenForm%>
 <%cboControlType%> <%lnkAddControl%><hr>
 <%
   Dim Control 
   if isarray(Controls) then
     For each Control in Controls 
        Control.Render() 
        Response.Write "<BR>" & vbnewline
     Next
   
   end if
 %>
<%Page.CloseForm%>
<hr>
</body>
</html>

<%
'### Controls
Dim cboControlType
Dim lnkAddControl

'### Variables

Dim Controls
Dim ControlNames
Dim ControlTypes
'

Function Page_Init() 
'	Page.DebugEnabled=true
	Set cboControlType = New_ServerDropDown("cboControlType")
	Set lnkAddControl  = New_ServerLinkButton("lnkAddControl")
End Function

Function Page_Controls_Init()
	cboControlType.Items.Append "","",True
	cboControlType.Items.Append "TextBox","TextBox",false
	cboControlType.Items.Append "Date","Date",false
	cboControlType.Items.Append "DropDown","DropDown",false
	lnkAddControl.Text = "Add Control"
End Function

Function Page_LoadViewState() 
    Dim hasControls,x,mx
    hasControls = Page.ViewState.GetValue("Types")<>""
    If hasControls Then
	   ControlTypes = Split(Page.ViewState.GetValue("Types"),"@")
	   ControlNames = Split(Page.ViewState.GetValue("Names"),"@")
	   mx = Ubound(ControlNames)
	   Redim Controls(mx)
	   
	   For x = 0 to mx
			Select Case ControlTypes(x)
				Case "TextBox"
					Set Controls(x) = New_ServerTextBox(ControlNames(x))
				Case "Date"
					Set Controls(x) = New_ServerDateTextBox(ControlNames(x))
				case "DropDown"
					Set Controls(x) = New_ServerDropDown(ControlNames(x))	
			End Select
	   Next
	End If
	
End Function

Function Page_SaveViewState()
   If IsArray(ControlNames) Then
   	  Page.ViewState.Add "Types",Join(ControlTypes,"@")
	  Page.ViewState.Add "Names",Join(ControlNames,"@")
   End If
End Function

'### Control PostBack Events

function lnkAddControl_OnClick() 
	Dim ctrl_name,ctrl_type,index 
	If cboControlType.Text = "" Then
	   Response.Write "Hey man!"
	   exit function
	End If
	
	if isarray(Controls) Then
	  Redim Preserve Controls(UBound(Controls) + 1 )
	  Redim Preserve ControlNames(UBound(ControlNames) + 1 )
	  Redim Preserve ControlTypes(UBound(ControlTypes) + 1 )
    else
      Redim Controls(0)
      Redim ControlNames(0)
      Redim ControlTypes(0)     
    end if
    index = Ubound(Controls)
    
    ctrl_name = "ctrl" & Ubound(Controls)
    ctrl_type = cboControlType.Text
    ControlNames(index) = ctrl_name
    ControlTypes(index) = ctrl_type
    
    Select Case cboControlType.Text
		Case "TextBox"
			Set Controls(index) = New_ServerTextBox(ctrl_name)
		Case "Date"
		    Set Controls(index) = New_ServerDateTextBox(ctrl_name)
		case "DropDown"
		    Set Controls(index) = New_ServerDropDown(ctrl_name)	
		    Controls(index).Items.Append "TEST1","TEST1",False
		    Controls(index).Items.Append "TEST2","TEST2",False
    End Select
    
end function

%>


