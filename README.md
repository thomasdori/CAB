Classic Asp Boilderplate (CAB)
================================

Addressed Issues
----------------
1. Website Defacement
	* All inputs will get sanitized in one place: Function InputHelper.getParameter()
	* All text outputs will get encoded in one place: Function OutputHelper.write()
	* All url outputs will get encoded in one place: Function OutputHelper.writeURL()
2. Stored XSS
	* see 1.
3. SQL Injection
	* see 1.
	* Only using stored procedures: Model.UserModel.getUserList()
4. DOS
	* see 1.
5. CSRF
	* Requests will only be sent via HTTP POST: CommunicationHandler.post() in script.js
	* Requests can only be used once because with every request the server will send a
  	  randomly generated token which will be sent back with the next request from the client:
  	  CsrfHelper.getParameter() and CsrfHelper.checkParameter() ond the server and CommunicationHandler.post() in script.js on the client
6. Clickjacking
	* Requests will only be sent via HTTP POST: CommunicationHandler.post() in script.js
	* Setting X-Frame option in ConfigurationHelper: Response.AddHeader "X-FRAME-OPTIONS", "DENY"
	* Adding JavaScript for browser which do not support X-Frame option: <antiClickjack> + related script block below
7. Information Leakage - Application Error
8. Information Leakage - Database Error
9. Content Tampering
10. Cookies not Marked HttpsOnly


Issues not addressed in CAB
---------------------------
1. Malicious File Upload
   Will not be addressed
2. Insecure Cryptographic Storage
   Could be addressed by encrypting email addresses but can not be tested with black box tests
3. Session ID remains the same after login
   Call Session.Abandon() after successfull login
4. Company Password/User/Password Enumeration
   Display only one error message
5. Weak Password Policy
   Implement password policy
6. Bruteforce Possible on Login
   Implement RECAPTCHA
7. Email Address Disclosure
   Don't display email addresses with @
8. Web Server Version Disclosure from HTTP Header
   Remove header - IIS setting
9. HTTPS only
   Server setting
10. Account Lockout
   Implement logic - after 10 unsuccessful attempts set account to inactive


Open Issues
-----------
* tests
* controller impl
* model impl
* dhtmlx
* error handling - no asp/server/sql messages
* httpsonly cookie
* doku
	dass am server css files interpretiert werden müssen
	classen dokumentieren
