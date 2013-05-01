<%
' Original Idea FROM:
' Barcode Generator (c) 2003, Lake Avenue Church, www.lakeave.org
'
' This program will generate the simpler Code39 (or Code 3-for-9) barcodes.
' You'll find these barcodes on staff and student ID badges, video rental
' cards and so on.
' Code39 works with alternating bands of black and white. Each character
' is represented by 7 bands. The band is either wide or narrow. Inbetween
' each character there is a band of white the same width as the narrow band.
' If you find the code useful, I'd love to hear from you. James Lamb, jamesl@lakeave.org
'
' Adapted for CLASP by Christian C.
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server BarCode
	Public Function New_ServerBarCode(name) 
		Set New_ServerBarCode = New ServerBarCode
			New_ServerBarCode.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerBarCode"

	Redim BarCodePtrn(44)
	BarCodePtrn(1)="1wnnwnnnnw"
	BarCodePtrn(2)="2nnwwnnnnw"
	BarCodePtrn(3)="3wnwwnnnnn"
	BarCodePtrn(4)="4nnnwwnnnw"
	BarCodePtrn(5)="5wnnwwnnnn"
	BarCodePtrn(6)="6nnwwwnnnn"
	BarCodePtrn(7)="7nnnwnnwnw"
	BarCodePtrn(8)="8wnnwnnwnn"
	BarCodePtrn(9)="9nnwwnnwnn"
	BarCodePtrn(10)="0nnnwwnwnn"
	BarCodePtrn(11)="Awnnnnwnnw"
	BarCodePtrn(12)="Bnnwnnwnnw"
	BarCodePtrn(13)="Cwnwnnwnnn"
	BarCodePtrn(14)="Dnnnnwwnnw"
	BarCodePtrn(15)="Ewnnnwwnnn"
	BarCodePtrn(16)="Fnnwnwwnnn"
	BarCodePtrn(17)="Gnnnnnwwnw"
	BarCodePtrn(18)="Hwnnnnwwnn"
	BarCodePtrn(19)="Innwnnwwnn"
	BarCodePtrn(20)="Jnnnnwwwnn"
	BarCodePtrn(21)="Kwnnnnnnww"
	BarCodePtrn(22)="Lnnwnnnnww"
	BarCodePtrn(23)="Mwnwnnnnwn"
	BarCodePtrn(24)="Nnnnnwnnww"
	BarCodePtrn(25)="Ownnnwnnwn"
	BarCodePtrn(26)="Pnnwnwnnwn"
	BarCodePtrn(27)="Qnnnnnnwww"
	BarCodePtrn(28)="Rwnnnnnwwn"
	BarCodePtrn(29)="Snnwnnnwwn"
	BarCodePtrn(30)="Tnnnnwnwwn"
	BarCodePtrn(31)="Uwwnnnnnnw"
	BarCodePtrn(32)="Vnwwnnnnnw"
	BarCodePtrn(33)="Wwwwnnnnnn"
	BarCodePtrn(34)="Xnwnnwnnnw"
	BarCodePtrn(35)="Ywwnnwnnnn"
	BarCodePtrn(36)="Znwwnwnnnn"
	BarCodePtrn(37)="-nwnnnnwnw"
	BarCodePtrn(38)=".wwnnnnwnn"
	BarCodePtrn(39)=" nwwnnnwnn"
	BarCodePtrn(40)="*nwnnwnwnn"
	BarCodePtrn(41)="$nwnwnwnnn"
	BarCodePtrn(42)="/nwnwnnnwn"
	BarCodePtrn(43)="+nwnnnwnwn"
	BarCodePtrn(44)="%nnnwnwnwn"

	 Class ServerBarCode
		Dim Control
		Dim Text
		Dim Narrow
		Dim Height
				
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Text    = ""	
			Narrow  = 0
			Height  = 10
	   End Sub
	   
	   	Public Function ReadProperties(bag)			
			Text = bag.Read("T")
			Narrow = CInt(bag.Read("N"))
			Height = CInt(bar.Read("H"))
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "T",Text
			bag.Write "N",Narrow
			bag.Write "H",Height
		End Function
	   
	   Public Function HandleClientEvent(e)
	   End Function			
	   
	   Public Function SetValueFromDataSource(value)
			Text = value
	   End Function
	   
	   Public Default Function Render()
			
			 Dim varStart	 
			 Dim strBarCode
			 Dim strConv
			 Dim t,s
			 Dim b			 	 
			 Dim sChar
			 Dim sImageBlack1
			 Dim sImageBlack2
			 Dim sImageClear1
			 Dim sImageClear2
			 
			 
			 If Control.Visible = False Then
				Exit Function
			 End If
			 
			 varStart = Now
			
			 strBarCode="*" & UCASE(Text) & "*"
			 strConv=""

			For t=1 To Len(strBarCode)
				For s=1 to 44
					If Mid(strBarCode,t,1) = Left(BarCodePtrn(s),1) Then
						strConv = strConv & Right(BarCodePtrn(s),9) & "s"
					End if
				Next
			Next
			
			sImageBlack1 = "<img src='" & SCRIPT_LIBRARY_PATH + "barcode/images/shim_black.gif' width=" & Narrow & " height=" & Height & ">"
			sImageBlack2 = "<img src='" & SCRIPT_LIBRARY_PATH + "barcode/images/shim_black.gif' width=" & Narrow*2 & " height=" & Height & ">"
			sImageClear1 = "<img src='" & SCRIPT_LIBRARY_PATH + "barcode/images/shim.gif' width=" & Narrow & " height=" & Height & ">"
			sImageClear2 = "<img src='" & SCRIPT_LIBRARY_PATH + "barcode/images/shim.gif' width=" & Narrow*2 & " height=" & Height & ">"															
			
			
		'	Response.Write "<BR>" & strConv & "<BR>"
		'	Response.Write "<BR>" & Ucase(Text) & "<BR>"
			
			For t=1 to Len(strConv)
				sChar = Mid(strConv,t,1)
				Select Case sChar
					Case "n","w"		
						If b=1 Then 
							Response.Write IIF(sChar = "n",sImageBlack1,sImageBlack2)
						ElseIf b = 0 Then
							Response.Write IIF(sChar = "n",sImageClear1,sImageClear2)
						End If
						b=b+1
						If b=2 Then 
							b=0
						End If
					Case "s"
						Response.Write sImageClear1
						b=1
					Case Else
						Response.Write "WHAT!"
				End Select
				
			Next
			Page.TraceRender varStart,Now,Control.Name
		
		End Function

	End Class

%>