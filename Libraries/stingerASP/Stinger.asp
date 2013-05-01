<%
'
 ' Stinger is an HTTP Request Validation Engine
 ' Originally developed by Aspect Security, Inc.
 '
 ' This library is free software you can redistribute it and/or
 ' modify it under the terms of the GNU Lesser General Public
 ' License as published by the Free Software Foundation either
 ' version 2.1 of the License, or (at your option) any later version.
 '
 ' This library is distributed in the hope that it will be useful,
 ' but WITHOUT ANY WARRANTY without even the implied warranty of
 ' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 ' Lesser General Public License for more details.
 '
 ' You should have received a copy of the GNU Lesser General Public
 ' License along with this library if not, write to the Free Software
 ' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 '
 '/

''
 ' Stinger provides the public API for the Stinger package and can be used to create RuleSets, apply them
 ' to HttpServletRequests, and handle the errors that are uncovered.  Essentially, Stinger models
 '  the validation requirements for an HTTP request. The Stinger singleton contains a list of
 ' RuleSets that may or may not apply to a particular HTTP request. Stinger also provides methods
 ' to generate the client side JavaScript to insert in a web page and format an error list into HTML.
 '
 ' @author jeff.williams@aspectsecurity.com
 '/

class Stinger

    ''
     ' Hide the constructor to force use of the getInstance() singleton method.
     '/
  Private Sub Class_Initialize()

  End Sub

  Private Sub Class_Terminate()

  End Sub

    ''
     ' Returns a JavaScript suitable for inserting in a web page that will validate the parameters in
     ' a form. The script displays an alert box containing ALL the errors found in the form. The
     ' script uses the exact same regular expressions as the server side validation. However, the
     ' JavaScript alone is easily bypassed and does not provide ANY security protection. It's only
     ' benefit is in ease of use for users with a slow Internet connection.
     '
     ' @param form the name of the form that the script will access to get parameters
     ' @param target the action URI targeted by the form
     '
     ' @return a String containing a JavaScript that will validate input on the client side and
     ' create a dialog box with the results.
     '/
    Public Function getJavaScript(form)
        Dim Rules, Rule, ParameterRules, i, n
        Set Rules = Server.CreateObject("MSXML2.DOMDocument")

        Dim scriptPart1, scriptPart2

        If Not Rules.Load(GetRulesPath) Then
          'do we want to display this information to user?
          'Can we afford validation not to occur due to SVDL file syntax errors
          'getJavaScript = "<script>alert('unable to load validation file')</script>"
        Else
            i = 0
            n = 0
            'Improved: the whole javascript was not working due to lack of semicolons???
            scriptPart1 = "<SCRIPT>" + vbcrlf
            scriptPart2 = "function validate() { msg='Errors found:\n'; err=0" + vbcrlf

            set ParameterRules = Rules.DocumentElement.SelectNodes ("rule[paramType='parameter']")
            For Each Rule In ParameterRules
                Dim name
                name = Rule.childNodes(0).NodeTypedValue

                ' only test if rule is a parameter and is not set to ignore malformed
                If UCase(Rule.selectSingleNode("malformedAction").NodeTypedValue) = "IGNORE" Then
                  Continue
                End If

                ' set up the regular expression
                Dim regex
                regex = Rule.selectsingleNode("regex").NodeTypedValue
                scriptPart1 = scriptPart1 + "regex" + CStr(n) + "=/" + regex + "/;" + vbcrlf

                Dim message, messageNode
                set messageNode = Rule.SelectSingleNode("malformedMessage")
                if not messageNode is nothing then
                  message = messageNode.NodeTypedValue
                else 'Improved: This message was empty, why not telling something if field is malformed
                  message = "Enter valid information on " & Server.HTMLEncode(name) & " field"
                end if

                ' if the parameter is not required
                'Improved: Looks like validations were swapped, value attribute was checked when ignore???
                If UCase(Rule.selectSingleNode("missingAction").NodeTypedValue) = "IGNORE" Then
                    'Improved: Why validate somthing you do not expect to be there??? I am commenting this code by know
                    'scriptPart2 = scriptPart2 + "if (!regex" + cstr(n) + ".test(" + form + "." + name + ".value)) {err+=1; msg+='\n  " + message + "'}" + vbcrlf
                    'n = n + 1
                Else                 ' if the parameter is required
                    'Improved: check if field is present correctly not just with .value attribute
                    scriptPart2 = scriptPart2 + "if ( " + form + "." + name + " == undefined || " + form + "." + name + ".value.length==0" + " || !regex" + cstr(n) + ".test(" + form + "." + name + ".value) ) {err+=1; msg+='\n  " + message + "'}" + vbcrlf
                    n = n + 1
                End If
            Next

            scriptPart2 = scriptPart2 + "if ( err > 0 ) alert(msg); else form.submit();"
            scriptPart2 = scriptPart2 + "} </SCRIPT>"
        End If
        getJavaScript = scriptPart1 + scriptPart2
    End Function

    ''
     ' Formats a list of errors into an HTML list for easy displaying
     '/
    Public Function format(errors)
        Dim sb, i
        sb = sb + vbcrlf
        For i = 1 To errors.Count
            sb = sb + "<DIV> - " + errors(i) + "</DIV>" + vbcrlf
        Next
        format = sb
    End Function

    Public Function GetRulesPath ()
      dim ScriptFullPath, ScriptPath, ScriptFile
      ScriptFullPath = Server.MapPath(Request.Servervariables("SCRIPT_NAME"))
      ScriptPath = left(ScriptFullPath, instrrev(ScriptFullPath, "\"))
      ScriptFile = mid(ScriptFullPath, len(ScriptPath)+1)
      GetRulesPath = ScriptPath + "rules\" + left(ScriptFile, len(ScriptFile)-4) + ".svdl.asp"
    End Function
    ''
     ' Validates the HTTP request against the rulesets in Stinger.
     '
     ' @param request the HTTP request to be validated.
     '
     ' @return a list of ValidationErrors detailing any problems found
     '/
    Public Function validate()
      Dim Rules, Rule, errors
      Set Rules = Server.CreateObject("MSXML2.DOMDocument")
      Set errors = Server.CreateObject("Scripting.Dictionary")
      If Not Rules.Load(GetRulesPath) Then
        'do we want to display this information to user?
        'Can we afford validation not to occur due to SVDL file syntax errors
      Else
        set errors = validateExtra(request, errors, Rules)
        set errors = validateMissing(request, errors, Rules)
        set errors = validateMalformed(request, errors, Rules)
      End If
      set validate = errors
    End Function

    ''
     ' Validate the request for extra parts using the list of rulesets supplied
     '
     ' @throws FatalValidationException
     '/
    Public Function validateExtra(request, ByRef errors, ByRef Rules)
      Dim FilteredCookieRules, FilteredHeaderRules, FilteredParamRules, ExtraCookieActionNode, ExtraCookieAction, ExtraHeaderActionNode, ExtraHeaderAction, ExtraParamAction, ExtraParamActionNode
      Dim cookie, header, Parameter
      'check cookies

      Set ExtraCookieActionNode = Rules.documentElement.SelectSingleNode("extraCookieAction")
      ExtraCookieAction = "FATAL"
      if not ExtraCookieActionNode is nothing then
        ExtraCookieAction = ExtraCookieActionNode.NodeTypedValue
      end if
      If UCASE(ExtraCookieAction) <> "IGNORE" Then
        for each cookie in request.cookies
          set FilteredCookieRules = Rules.documentElement.selectNodes("rule[paramType='cookie' and name='" + cookie + "']")
          If FilteredCookieRules.Length = 0 Then
            If UCase(ExtraCookieAction) = "CONTINUE" Then
              errors.Add errors.count+1, "The [" + Server.HTMLEncode(cookie) + "] cookie was unexpected"
            Else
              Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(cookie) + "] cookie was unexpected"
            End If
          End If
        Next
      End If

      set ExtraHeaderActionNode = Rules.documentElement.SelectSingleNode("extraHeaderAction")
      ExtraHeaderAction = "FATAL"
      if not ExtraHeaderActionNode is nothing then
        ExtraHeaderAction = ExtraHeaderActionNode.NodeTypedValue
      end if
      If UCase(ExtraHeaderAction) <> "IGNORE" Then
        for each header in request.ServerVariables
          set FilteredHeaderRules = Rules.documentElement.selectNodes("rule[paramType='header' and name='" + header + "']")
          If FilteredHeaderRules.Length = 0 Then
            If UCase(ExtraHeaderAction) = "CONTINUE" Then
              errors.Add errors.count+1, "The [" + Server.HTMLEncode(header) + "] header was unexpected"
            Else
              Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(header) + "] header was unexpected"
            End If
          End If
        Next
      End If

      Set ExtraParamActionNode = Rules.documentElement.SelectSingleNode("extraParameterAction")
      ExtraParamAction = "FATAL"
      if not ExtraParamActionNode is nothing then
        ExtraParamAction = ExtraParamActionNode.NodeTypedValue
      end if
      If UCase(ExtraParamAction) <> "IGNORE" Then
        for each parameter in request.QueryString
          set FilteredParamRules = Rules.documentElement.selectNodes("rule[paramType='parameter' and name='" + Parameter + "']")
          If FilteredParamRules.Length = 0 Then
            If UCase(ExtraParamAction) = "CONTINUE" Then
              errors.Add errors.count+1, "The [" + Server.HTMLEncode(Parameter) + "] parameter was unexpected"
            Else
              Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(Parameter) + "] parameter was unexpected"
            End If
          End If
        Next
     For each parameter in request.Form
        set FilteredParamRules = Rules.documentElement.selectNodes("rule[paramType='parameter' and name='" + Parameter + "']")
        If FilteredParamRules.Length = 0 Then
          If UCase(ExtraParamAction) = "CONTINUE" Then
            errors.Add errors.count+1, "The [" + Server.HTMLEncode(Parameter) + "] parameter was unexpected"
          Else
            Err.Raise 1, "Stinger", "The [" + Server.HTMLEncode(Parameter) + "] parameter was unexpected"
          End If
        End If
      Next
    End If

    Set validateExtra = errors
  End Function


    ''
     ' Validate the request for malformed parts using the list of rulesets supplied
     '
     ' @throws FatalValidationException
     '/
    Private Function validateMalformed(request, ByRef errors, ByRef Rules)
      Dim AllRules, Rule, RuleName, RequestValue
      Dim MalformedActionNode, MalformedAction, MalformedMessageNode, MalformedMessage
      Set AllRules = Rules.documentElement.selectNodes("rule")
      For Each Rule In AllRules
        RuleName = Rule.childNodes(0).nodeTypedValue
        Set MalformedActionNode = Rule.selectSingleNode("malformedAction")
        MalformedAction = "FATAL"
        if not MalformedActionNode is nothing then
          MalformedAction = MalformedActionNode.NodeTypedValue
        End if
        if UCase(MalformedAction) <> "IGNORE" then
          RequestValue = request(RuleName)
          If VarType(RequestValue) <> 0 Then 'data exists for name element in rule
            Dim regex
            Set regex = New RegExp
            regex.Pattern = Rule.childNodes(2).nodeTypedValue 'Regex item
            regex.IgnoreCase = True   ' Set case insensitivity.
            regex.Global = True   ' Set global applicability.
            If regex.Test(RequestValue) Then
              Set validateMalformed = errors
            Else
              MalformedMessage = "Malformed " + Rule.childNodes(1).nodeTypedValue + " '" + RuleName + "' = '" + RequestValue + "'"
              Set MalformedMessageNode = Rule.selectSingleNode("malformedMessage")
              if not MalformedMessageNode is nothing then
                MalformedMessage = MalformedMessageNode.NodeTypedValue
              end if

              If UCase(MalformedAction) = "FATAL" Then
                Err.Raise 1, "Stinger", MalformedMessage
              Else
                errors.Add errors.Count + 1, MalformedMessage
              End If
            End If
          End If
        End If
      Next
      Set validateMalformed = errors
    End Function

    ''
     ' Validate the request for missing parts using the list of rulesets supplied
     '
     ' @throws FatalValidationException
     '/
    Private Function validateMissing(request, ByRef errors, ByRef Rules)
      Dim AllRules, Rule, RuleName, RequestValue, MissingActionNode, MissingAction, MissingMessageNode, MissingMessage
      Set AllRules = Rules.documentElement.selectNodes("rule")
      For Each Rule In AllRules
        RuleName = Rule.childNodes(0).nodeTypedValue
        Set MissingActionNode = Rule.selectSingleNode("missingAction")
        MissingAction = "FATAL"
        if not MissingActionNode is nothing then
          MissingAction = MissingActionNode.NodeTypedValue
        End if
        if UCase(MissingAction) <> "IGNORE" then
          RequestValue = request(RuleName)
          If VarType(RequestValue) = 0 Then 'data exists for name element in rule
            MissingMessage = "Missing " + Rule.childNodes(1).nodeTypedValue + " '" + RuleName + "'"
            Set MissingMessageNode = Rule.selectSingleNode("missingMessage")
            if not MissingMessageNode is nothing then
              MissingMessage = MissingMessageNode.NodeTypedValue
            end if
            If UCase(MissingAction) = "FATAL" Then
              Err.Raise 1, "Stinger", MissingMessage
            Else
              errors.Add errors.Count + 1, MissingMessage
            End If
          End If
        End if
      Next
      Set validateMissing = errors
    End Function
End Class
%>