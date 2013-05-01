<!--#Include File = "WebControl.asp"-->
<%
Dim PageController
Dim CurrentUser

Set	CurrentUser		= New cApplicationUser
Set PageController	= New cPageControllerClass

'PageController.FolderSecurity.Add FolderName,  RequiredAuthentication,RequiresAuthorization,AuthorizedRoles
'use FolderNAme.Split("/")  and loop to get the high ranking match
'file sec overrides folder sec
Class cPageControllerClass
			
	Public  AuthorizedRoles
	Public  RequiresAuthentication
	Public  RequiresAuthorization
	Public  PageTitle			
	Public  ParanoidLevel
	Public  PageCSSFiles	
	Public  PageJavaScriptFiles
	Public  BodyAttributes	
	Public 	HideHeader
	Public 	HideFooter
		
	Public  PublicHomeURL
	Public  SecuredHomeURL

	'Public	FolderSecurity
	 	
	Private Sub Class_Initialize()			
		PageTitle		= "No Title"				
		RequiresAuthentication	= True
		RequiresAuthorization   = True
		ParanoidLevel	= 0 '0 Optimistic, 1 Medium, 3 VERY Paranoid
		HideHeader = False
		HideFooter = False
	End Sub
			
	Private Sub Class_Terminate()
	End Sub
	
	Private Function HandleParanoidLevel()
		If ParanoidLevel > 0 Then
			If Left(Request.ServerVariables("HTTP_REFERER"),LEN("HTTP:\\LOCALHOST"))<> "HTTP:\\LOCALHOST" THEN
				Response.Redirect "CannotPostFromAnotherWS.html"
				Response.End 
			End If
		End If
	End Function
	
	Public Function OpenPage()
		Call  ExecuteEventFunction("Page_Configure")
		Call  HandleParanoidLevel()
		Response.Write vbNewLine & "<HTML>"
		Response.Write vbNewLine & "<HEAD>"
			Call  Page.Execute()
			Response.Write vbNewLine & "<TITLE>" & PageTitle & "</TITLE>"
			Call  ExecuteEventFunction("Page_RenderHeadTag")		
			Call  RenderCollectionByTemplate("<LINK REL='STYLESHEET' TYPE='text/css' HREF='{0}'>",PageCSSFiles)
			Call  RenderCollectionByTemplate("<SCRIPT LANGUAGE='JavaScript' SRC='{0}'></SCRIPT>",PageJavaScriptFiles)
		Response.Write vbNewLine & "</HEAD>"
		Response.Write vbNewLine & "<BODY ID='CLASPBody' " & BodyAttributes & ">"
		Call  Page.OpenForm()
			If Not HideHeader Then Call  ExecuteEventFunction("Page_RenderHeader")
				Call  ExecuteEventFunction("Page_RenderForm")
			If Not HideFooter Then Call  ExecuteEventFunction("Page_RenderFooter")
		Call  Page.CloseForm()
	End Function

	Public Function ClosePage()
		Response.Write vbNewLine & "</BODY>"
		Response.Write vbNewLine & "</HTML>"
	End Function

	Public Function RenderPage()
		Call OpenPage()
		Call ClosePage()
	End Function

	Public Function GoToHomePage()
		If CurrentUser.IsAuthenticated Then
			Response.Redirect PageController.SecuredHomeURL
		Else
			Response.Redirect PageController.PublicHomeURL
		End If		
	End Function

	Private Function RenderCollectionByTemplate(Template,Values)
		Dim value
		If IsArray(Values) Then
			For Each value in Values
				Response.Write Replace(Template,"{0}",value,1,1)
			Next
		Else
			If Values <> "" Then
				Response.Write Replace(Template,"{0}",values,1,1)
			End If
		End If	
	End Function	

End Class

'USER & SECURITY CLASS

Class cApplicationUser
	
	Private mAuthToken
	Private mUserSessionPrefix
	Private mUserSecLocation '(0)CLIENT/(1)SERVER
	
	Public  NoAccessURL
	Public  LoginURL
	
	Private mUserID
	Private mFirstName
	Private mLastName
	Private mEmail
	Private mRoles

	
	Private Sub Class_Initialize()		
		
		mUserSessionPrefix	= "_USR_"
		mAuthToken			= mUserSessionPrefix & "__AUTHORIZED"		
		mUserSecLocation	= 1
		
		mUserID		=	""
		mFirstName	=	""
		mLastName	=	""
		mEmail		=	""
		mRoles		=   Array("")
		
		NoAccessURL = IIF(Application.Contents("CLASP_NOACCESS_URL")<>"",Application.Contents("CLASP_NOACCESS_URL"),"NoAccess.asp")
		LoginURL    = IIF(Application.Contents("CLASP_LOGIN_URL")<>"",Application.Contents("CLASP_LOGIN_URL"),"Login.asp")		
		
		If Session.Contents.Item(mAuthToken) <> "" Then
			Call LoadUser()
		End If		
		
	End Sub
			
	Private Sub Class_Terminate()		
	End Sub

	Public Property Get UserID()
		UserID = Session.Contents.Item(mUserSessionPrefix & "UserID") 
	End Property
	
	Public Property Get FirstName()
		FirstName = Session.Contents.Item(mUserSessionPrefix & "FirstName") 
	End Property

	Public Property Get LastName()
		LastName = Session.Contents.Item(mUserSessionPrefix & "LastName") 
	End Property

	Public Property Get Email()
		Email = Session.Contents.Item(mUserSessionPrefix & "Email") 
	End Property

	Public Property Let Roles(arrRoles)
		mRoles = arrRoles
		Session.Contents.Item(mUserSessionPrefix & "Roles") =Join(mRoles,",")
	End Property

	Public Property Get Roles()
		Roles = mRoles 
	End Property

	Private Sub LoadUser()
		mUserID			= Session.Contents.Item(mUserSessionPrefix & "UserID") 
		mFirstName		= Session.Contents.Item(mUserSessionPrefix & "FirstName") 
		mLastName		= Session.Contents.Item(mUserSessionPrefix & "LastName") 
		mEmail			= Session.Contents.Item(mUserSessionPrefix & "Email") 
		mRoles			= Split(Session.Contents.Item(mUserSessionPrefix & "Roles"),",") 		
	End Sub
	
	Private Sub SaveUser()
		Session.Contents.Item(mUserSessionPrefix & "UserID")	=	mUserID
		Session.Contents.Item(mUserSessionPrefix & "FirstName") =	mFirstName
		Session.Contents.Item(mUserSessionPrefix & "LastName")	=	mLastName
		Session.Contents.Item(mUserSessionPrefix & "Email")		=	mEmail
		Session.Contents.Item(mUserSessionPrefix & "Roles")		=	Join(mRoles,",")
	End Sub	

	Private Sub ClearUser()
		Session.Contents.Item(mUserSessionPrefix & "UserID")	=	""
		Session.Contents.Item(mUserSessionPrefix & "FirstName") =	""
		Session.Contents.Item(mUserSessionPrefix & "LastName")	=	""
		Session.Contents.Item(mUserSessionPrefix & "Email")		=	""
		Session.Contents.Item(mUserSessionPrefix & "Roles")		=	Join("",",")
		Session.Contents.Item(mAuthToken)						=	"NO"
	End Sub	

	Public Function RedirectFromLoginPage(UserName,FirstName,LastName,Email,Roles)
		Dim sReturnURL
		
		sReturnURL = Trim(Session.Contents.Item(mUserSessionPrefix  & "RedirectTo"))
		Authenticate UserName,FirstName,LastName,Email,Roles
		If  sReturnURL <> "" Then
			Response.Redirect sReturnURL
			Response.End
			Exit Function
		End If
		PageController.GoToHomePage		
		
	End Function
	Public Function Authenticate(UserName,FirstName,LastName,Email,Roles)
		
		Session.Contents.Item(mAuthToken) = "YES"		
		mUserID		= UserName
		mFirstName	= FirstName
		mLastName	= LastName
		mEmail		= Email
		If IsArray(Roles) Then
			mRoles = Roles
		Else
			mRoles = Split(Roles,",")
		End If		
		Call SaveUser()
	End Function
			
	Public Property Get IsAuthenticated
		IsAuthenticated = (Session.Contents.Item(mAuthToken) = "YES")
	End Property
	
	Public Function IsInRole(RoleName)
		Dim role
		IsInRole = False
		
		If IsEmpty(PageController.AuthorizedRoles) Then
			IsInRole = True
			Exit Function
		End If
		
		For Each role in mRoles
			If role = RoleName Then
				IsInRole = True
				Exit For
			End If
		Next
	End Function
	
	Public Function IsInRoles(Roles)
		Dim PageRole,UserRole
		IsInRoles = False
				
		For Each UserRole in mRoles
			For Each PageRole in Roles
				If UserRole = PageRole Then
					IsInRoles	= True
					Exit For
				End If
			Next
		Next
	End Function
	
	Public Function LogOut()
	    Session.Contents.Item(mAuthToken) = "NO"		
		Session.Abandon		
		Response.Redirect LoginURL
	End Function

	Public Function GoToLoginPage()
		Session.Contents.Item(mUserSessionPrefix  & "RedirectTo") = Request.ServerVariables("PATH_INFO") & IIF(Request.ServerVariables("QUERY_STRING")<>"","?" & Request.ServerVariables("QUERY_STRING"),"")
		Response.Redirect LoginURL
		Response.End	
	End Function

	Public Function Authorize()
		Dim bolAuthorized		
				
		If IsArray(PageController.AuthorizedRoles) Then
			bolAuthorized = IsInRoles(PageController.AuthorizedRoles)
		Else
			bolAuthorized = IsInRole(PageController.AuthorizedRoles)
		End If
		
		If Not bolAuthorized Then
			Response.Redirect NoAccessURL
			Response.End
		End If
		
	End Function
End Class

':: CUSTOM PAGE EVENTS THAT ARE GLOBAL TO ALL PAGES
Public Function Page_LogOut(e)
	Call CurrentUser.LogOut()
End Function

':: PAGE EVENTS/SECURITY
Public	Function Page_Authenticate_Request()
	'Make sure that the login page does not require authentication!
	If PageController.RequiresAuthentication Then
		If Not CurrentUser.IsAuthenticated Then
			CurrentUser.GoToLoginPage
		End If	
	End If
End	Function

Public	Function Page_Authorize_Request()	
	If PageController.RequiresAuthentication And PageController.RequiresAuthorization Then
		Call CurrentUser.Authorize()
	End If
End Function
%>	