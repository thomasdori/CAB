<%
	'Dependencies: EncodingHelper.asp

	Dim cookie
	Set cookie = new CookieHelper

	Dim safeCookies
	Set safeCookies = Server.CreateObject("Scripting.Dictionary")
	Dim safeCookieParams
	Set safeCookieParams = Server.CreateObject("Scripting.Dictionary")

	Class CookieHelper
		Function read(key)
			read = encoder.encode(Request.Cookies(key))
		End Function

		Sub write(key, value)
			Call SetSafeCookieOption(key, "expires", Now)
			Call SetSafeCookieOption(key, "secure", true)
			Call SetSafeCookieOption(key, "httponly", true)
			Call SetSafeCookieOption(key, "", value)
			WriteSafeCookies
		End Sub

		' Source: http://blog.embrodesign.com/2012/10/asp-how-to-support-httponly-cookies/
		' Store a cookie name/value pair (must call WriteSafeCookies to flush to header)
		Function SetSafeCookie(cookiename, cookiefield, cookievalue)
			Dim safeCookie
			' Create a cookie entry if one doesn't exist
			If safeCookies.Exists(cookiename) Then
				Set safeCookie = safeCookies.Item(cookiename)
				safeCookie.Add cookiefield, cookievalue
				' Update the collection
				Set safeCookies.Item(cookiename) = safeCookie
			Else
				Set safeCookie=Server.CreateObject("Scripting.Dictionary")
				safeCookie.Add cookiefield, cookievalue
				' Add it to the collection of cookies
				safeCookies.Add cookiename, safeCookie
			End If
		End Function

		' Store a cookie parameter/setting (such as expires, path, domain, httponly, secure)
		Function SetSafeCookieOption(cookiename, param, paramvalue)
			Dim safeCookieParam
			isvalid=false

			SELECT CASE LCase(param)
				CASE "expires"
					isvalid=true
				CASE "path"
					isvalid=true
				CASE "domain"
					isvalid=true
				CASE "httponly"
					isvalid=true
				CASE "secure"
					isvalid=true
			END SELECT

			If isvalid Then
				' Create a cookie entry if one doesn't exist
				If safeCookieParams.Exists(cookiename) Then
					Set safeCookieParam = safeCookieParams.Item(cookiename)
					safeCookieParam.Add LCase(param), paramvalue
					' Update the collection
					Set safeCookieParams.Item(cookiename) = safeCookieParam
				Else
					Set safeCookieParam=Server.CreateObject("Scripting.Dictionary")
					safeCookieParam.Add LCase(param), paramvalue
					' Add it to the collection of cookies
					safeCookieParams.Add cookiename, safeCookieParam
				End If
			End If
		End Function

		' Write all cookies to the header stream
		Function WriteSafeCookies
			ishttponly = True    'Default HTTPOnly cookie
			issecure = False    'Default insecure cookie
			'Quick hack to support cookies that only specify parameters
			If safeCookies.Count = 0 And safeCookieParams.Count > 0 Then SetSafeCookie safeCookieParams.Keys()(0), "", ""

				For Each cookiekey In safeCookies.Keys
					cookiestr = ""
					Dim cookie
					Set cookie = safeCookies(cookiekey)
					' Do cookie values
					For Each datakey In cookie.Keys
						If Trim(datakey) <> "" And Trim(cookie(datakey)) <> "" Then
							cookiestr = cookiestr & Server.URLEncode(datakey) & "=" & Server.URLEncode(cookie(datakey)) & "&"
						ElseIf Trim(datakey) = "" And Trim(cookie(datakey)) <> "" Then
							cookiestr = cookiestr & Server.URLEncode(cookie(datakey)) & " "
						End If
					Next

					If Len(cookiestr) > 0 Then
						cookiestr = Left(cookiestr, Len(cookiestr) - 1)
					End If
					cookiestr = cookiestr & "; "
					' Do cookie parameters
					If safeCookieParams.Exists(cookiekey) Then
						For Each paramkey In safeCookieParams.Item(cookiekey).Keys
							Select Case LCase(paramkey)
							Case "expires"
								' Here we do a Timezone offset from the server registry
								currentDate = CDate(safeCookieParams.Item(cookiekey)(paramkey))
								cookiestr = cookiestr & "Expires" & "=" & ToGMTDateTime(currentDate) & "; "
							Case "path"
								cookiestr = cookiestr & "Path" & "=" & safeCookieParams.Item(cookiekey)(paramkey) & "; "
							Case "domain"
								cookiestr = cookiestr & "Domain" & "=" & safeCookieParams.Item(cookiekey)(paramkey) & "; "
							Case "secure"
								If safeCookieParams.Item(cookiekey)(paramkey) = false Then
								issecure = False
								Else
								issecure = True
								End If
							case "httponly"
								If safeCookieParams.Item(cookiekey)(paramkey) = false Then
									ishttponly = False
								Else
									ishttponly = True
								End If
							End Select

							If InStr(LCase(cookiestr), "path=") = 0 Then
								cookiestr = cookiestr & "Path=/; "
							End If
						Next
					End If

					cookiestr = cookiekey & "=" & cookiestr
					If issecure Then cookiestr = cookiestr & " Secure; "
					If ishttponly Then cookiestr = cookiestr & "HttpOnly; "

					'Remove padding and semicolon from end
					cookiestr = Left(cookiestr, Len(cookiestr) - 2)
					Response.AddHeader "Set-Cookie", cookiestr
					'Response.Write "Set-Cookie: " & cookiestr & "<br/>"

				Next
			Set safeCookies=Server.CreateObject("Scripting.Dictionary")
			Set safeCookieParams=Server.CreateObject("Scripting.Dictionary")
		End Function

		' Convert a date to GMT date time format
		Function ToGMTDateTime(currentDate)
			' Here we do a Timezone offset from the server registry, IIS/IUSR account must have access to registry for this to work (which is normal)
			Set oShell = CreateObject("WScript.Shell")
			atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
			offsetMin = oShell.RegRead(atb)
			currentDate = dateadd("n", offsetMin, currentDate)

			' Now the time is corrected for GMT, get the day of week in a 3 char format
			daystr = ""
			Select Case(Weekday(currentDate))
			Case 1
				daystr = "Sun"
			Case 2
				daystr = "Mon"
			Case 3
				daystr = "Tue"
			Case 4
				daystr = "Wed"
			Case 5
				daystr = "Thu"
			Case 6
				daystr = "Fri"
			Case 7
				daystr = "Sat"
			End Select

			ToGMTDateTime = daystr & ", " & Right("00" & Day(currentDate), 2) & "-" & MonthName(Month(currentDate), True) & "-" & Year(currentDate) & " " & Right("00" & Hour(currentDate), 2) & ":" & Right("00" & Minute(currentDate), 2) & ":" & Right("00" & Second(currentDate), 2) & " GMT"

		End Function
	End Class
%>