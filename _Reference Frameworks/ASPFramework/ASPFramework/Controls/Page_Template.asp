<!--#Include File = "Controls\WebControl.asp"        -->
<!--Include Any Extra Controls Here!-->
<%
':: STEP 1 - ABOVE, ADD ALL THE APPROPRIATE SERVER SIDE INCLUDES
%>	
<HTML>
<HEAD>
	<TITLE>Page Title</TITLE>
</HEAD>
<BODY>
<%
':: STEP 2 - EXECUTE THE PAGE AS SOON AS ALL THE INCLUDES HAVE BEEN ADDED TO THE PAGE
	Page.Execute
%>
<%':: STEP 3 - INCLUDE YOUR HTML HERE. YOU CAN ALSO ADD SERVER CONTROLS THAT DON'T RENDER ANY FORM INPUTS AS THEIR VALUES WOULD NOT BE POSTED %>
<%':: STEP 4 - USE OPENFORM TO OPEN THE FORM %>
<%Page.OpenForm%>
<%':: STEP 5 - INCLUDE YOUR SERVER CONTROLS, HTML AND ANY EXTRA HTML INPUT%>
<%Page.CloseForm%>
<%':: STEP 6 - INCLUDE CLOSEFORM TO CLOSE THE FORM %>
<%':: STEP 7 - INCLUDE MORE HTML HERE. YOU CAN ALSO ADD SERVER CONTROLS THAT DON'T RENDER ANY FORM INPUTS AS THEIR VALUES WOULD NOT BE POSTED %>
</BODY>
</HTML>

<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: STEP 8 - YOUR CODE BEHIND GOES HERE.                                      ::
'::          ALL PAGE VARIABLES SHOULD BE DECLARED HERE. ALL CODE FROM THIS   ::
'::          POINT ON SHOULD RESIDE INSIDE A FUNCTION/EVENT HANDLER           ::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::			   
':: 8.1  PAGE VARIABLES DECLARATION
':: i.e. Dim PageTitle
'::      Dim CurrentStatus
'::
':: 8.2  SERVER SIDE CONTROL DECLARATION
':: i.e. Dim txtFirstName
'::      Dim cmdSave
'::
':: 8.3 PAGE EVENTS GO HERE. IT IS SAFE TO REMOVE THE ONES YOU DON'T USE. ONLY
'::     Page_Init and  Page_Controls_Init almost always used.

Public Function Page_Authenticate_Request()
':: ANY CUSTOM AUTHENTICATION CODE. YOU MAY WANT TO MOVE THIS TO A SEPARATE INCLUDE FILE AND
':: ADD IT TO THE TOP OF THE PAGE
End Function
	
Public Function  Page_Authorize_Request()
':: ANY CUSTOM AUTHORIZATION CODE. YOU MAY WANT TO MOVE THIS TO A SEPARATE INCLUDE FILE AND
':: ADD IT TO THE TOP OF THE PAGE
End Function
	
Public Function Page_Init()
':: HERE YOU INITIALIZE ALL PAGE VARIABLES TO THEIR DEFAULT VALUES AND ALSO
':: INITIALIZE THE SERVER CONTROLS (i.e. PageTitle = "CLASP" or Set txtFirstName = New_ServerTextBox("txtFirstName")
End Function

Public Function Page_Controls_Init()						
':: HERE YOU SET THE PROPERTIES FOR THE CONTROLS THAT YOU WANT TO BE PERSISTED IN THE VIEWSTATE
':: OR ANY CODE THAT YOU WANT TO RUN ONLY THEN THE PAGE IS FIRST LOADED (Page.IsPostBack = False)
':: THIS EVENT OCCURS ONLY ONCE.
End Function
	
Public Function Page_LoadViewState()
':: HERE YOU CAN RESTORE VARIABLES THAT WERE ADDED TO THE VIEWSTATE
':: i.e. CurrentStatus = Page.ViewState.GetValue("CurrentStatus")
End Function
	
Public Function Page_Load()
':: AT THIS POINT ALL CONTROLS HAVE BEEN INITIALIZED, ALL VARIABLES SHOULD HAVE BEEN RESTORED
':: AND THE POSTBACK VALUES HAVE BEEN APPLIED. YOU CAN DO ANYTHING YOU WANT HERE THAT IS NOT NORMALLY
':: DEPENDING ON THE OUTCOME OF THE POSTBACK, SINCE IT IS EXECUTED RIGHT AFTER THIS STEP
':: NEXT TO THIS EVENT, ALL THE OnChanged EVENTS WILL BE TRIGGERED (i.e. txtFirstName_OnChanged, ONLY FOR THE CONTROLS YOU WANT TO
':: CAPTURE THIS EVENT) AND ALSO THE FUNCTION TO HANDLE THE POSTBACK WILL BE INVOKED (i.e. cmdSave_OnClick() )
End Function
		
Public Function Page_PreRender()
':: LAST MINUTE MANIPULATIONS COULD BE DONE HERE SINCE THE POSTBACK WAS ALREADY HANDLED. 
':: AT THIS POINT EVERYTHING SHOULD BE READY TO GO.
End Function

Public Function Page_SaveViewState()
':: HERE YOU CAN STORE VARIABLES THAT IN THE VIEWSTATE
':: i.e. Page.ViewState.Add "CurrentStatus",CurrentStatus
End Function

Public Function Page_Terminate()
':: DONE, THE PAGE WAS RENDERED. THERE IS NOTHING MUCH YOU CAN DO NOW!
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: STEP 9 - ADD YOUR POSTBACK EVENT HANDLERS HERE, SUCH AS cmdSave_OnClick   ::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: STEP 10 - ADD YOUR PAGE HELPER ROUTINES, SUCH AS LoadRecord(), ETC        ::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

%>
