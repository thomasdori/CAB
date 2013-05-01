Classic Asp Boilderplate (CAB)
================================



Description
-----------
The Classic Asp Boilerplate Procject (CAB) is an approach to solve common quality assurance issues like security, maintainability.
It also delivers an architecture outline that can be used for future projects.



Security Issues
-------------------------
1. Website Defacement
	* All inputs will get sanitized in one place: Function InputHelper.getParameter()
	* All text outputs will get encoded in one place: Function OutputHelper.write()
	* All url outputs will get encoded in one place: Function OutputHelper.writeURL()
	* <%=variable%> is not used. <% output.write(variable) %> is used instead.
	* Cookie values must not be shown in the UI

2. Stored XSS
	* see 1. Website Defacement

3. SQL Injection
	* see 1. Website Defacement
	* Only using stored procedures: Model.UserModel.getUserList()

4. DOS
	* see 1. Website Defacement

5. CSRF
	* Requests will only be sent via HTTP POST: CommunicationHandler.post() in script.js
	* Requests can only be used once because with every request the server will send a
  	  randomly generated token which will be sent back with the next request from the client:
  	  CsrfHelper.getParameter() and CsrfHelper.checkParameter() ond the server and CommunicationHandler.post() in script.js on the client

6. Clickjacking
	* Requests will only be sent via HTTP POST: CommunicationHandler.post() in script.js
	* Setting X-Frame option in ConfigurationHelper: Response.AddHeader "X-FRAME-OPTIONS", "DENY"
	* Adding JavaScript for browser which do not support X-Frame option: <antiClickjack> + related script block below

8. Information Leakage
	* In general errors should not be considered (On Error Resume Next in ConfigurationHelper.asp)
	* Application and database errors should not be delivered to the client

9. Content Tampering
	* Because URL parameter are not used (see 5. CSRF) they can not be displayed on the page.

10. Cookies not Marked HttpsOnly
	* All cookies are accessed via the CookieHelper class
	* Usage of cookies not using HttpOnly is prohibitted




Maintainability Issues
----------------------
_TODO_



Security Issues Not Addressed in CAB
---------------------------
1. Malicious File Upload
	* Will not be addressed

2. Insecure Cryptographic Storage
	* Fix: Could be addressed by encrypting email addresses but can not be tested with black box tests

7. Information Leakage - Application Error
	* Fix: Apply http://www.reedolsen.com/show-errors-for-classic-asp-pages-in-iis-6/

7. Information Leakage - Database Error
	* Fix: SERVER SETTING

3. Session ID remains the same after login
	* Fix: Call Session.Abandon() after successfull login

4. Company Password/User/Password Enumeration
	* Fix: Display only one error message

5. Weak Password Policy
	* Fix: Implement password policy

6. Bruteforce Possible on Login
	* Fix: Implement RECAPTCHA

7. Email Address Disclosure
	* Fix: Don't display email addresses with @

8. Web Server Version Disclosure from HTTP Header
	* Fix Remove header - SERVER SETTING

9. HTTPS only
	* Fix: SERVER SETTING

10. Account Lockout
	* Fix: Implement logic - e.g. after 10 unsuccessful attempts set account to inactive




Open Issues
-----------
### Important
* Tests: stinger

### Would be Nice
* Cast parameter in GetParameter
* javascript frameworks like dhtmlx + ActivityIndicator
* Error handling - no asp/server/sql messages
* Inline code documentation could be more extensive
