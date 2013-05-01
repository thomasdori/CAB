<%
Class StringBuilder
	Private strArray()
	Private intGrowRate
	Private intItemCount
	Private Sub Class_Initialize()		
		intGrowRate = 50
		intItemCount = 0
	End Sub
	
	Public Property Get GrowRate
		GrowRate = intGrowRate
	End Property

	Public Property Let GrowRate(value)
		 intGrowRate = value
	End Property
	
	Private Sub InitArray()
			Redim Preserve strArray(intGrowRate)
	End Sub
	
	Public Sub Append(str)
			
		If intItemCount = 0 Then
			Call InitArray
		ElseIf intItemCount > UBound(strArray) Then			
			Redim Preserve strArray(Ubound(strArray) + intGrowRate)
		End If
		
		strArray(intItemCount) = str
		
		intItemCount = intItemCount + 1
	
	End Sub	

	Public Function FindString(str)
		Dim x,mx
		mx = intItemCount - 1
		For x = 0 To mx
			If strArray(x) = str Then
				FindString = x
				Exit Function
			End If
		Next
		FindString = -1
	End Function
	
	Public Function ToString2(sep)
		If intItemCount = 0 Then
			ToString2 = ""
		Else
			Redim Preserve strArray(intItemCount)
			ToString2 = Join(strArray,sep)
		End If		
	End Function
		
	Public Default Function ToString()
		If intItemCount = 0 Then
			ToString = ""
		Else
			ToString = Join(strArray,"")
		End If		
	End Function

End Class
%>