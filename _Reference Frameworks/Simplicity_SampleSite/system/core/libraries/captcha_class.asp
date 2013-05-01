<%
' ASPCanvas 2.0.2, the revenge
' The object of this version is a complete overhaul of the backend image map
' by converting it all to ADO streams. This should yield a speed increase of
' around 15x the previous version.
' This version will also be a lot trimmer than the earlier version. A lot of 
' the unused stuff has been removed.
'
' Additions should include:
' Barcode production
' OCR-A OCR-B font
' Any other graphical drawing application I can think of.
Public Const PI = 3.14159265 ' Roughly

Class Canvas
	Public GlobalColourTable()

	Public BackgroundColourIndex
	Public ForegroundColourIndex
	Public TransparentColourIndex

	Public UseTransparency

	Public Comment

	Public ErrorsToIISLog
	Public ErrorsToResponse
	Public ErrorsHalt
	Public ErrorsToImage
	
	Public BMPWarnings
	
	Private m_objImage
	Private m_lWidth
	Private m_lHeight
	Private m_lPow
	Private m_lMask

	public property get Version()
		Version = "2.0.2"
	end property

	' The pixels are now ZERO based. All of a sudden, everything just got easier.
	' Hooray for streams!
	Public Property Let Pixel(ByVal lX,ByVal lY,ByVal lValue)
		if (lX < m_lWidth - 1 And lX >= 0) And (lY < m_lHeight - 1And lY >= 0) then
			m_objImage.Position = (CLng(lY) * m_lWidth) + CLng(lX)
			m_objImage.WriteText Chr(lValue)
		end if
	End Property

	Public Property Get Pixel(ByVal lX,ByVal lY)
		if (lX < m_lWidth - 1 And lX >= 0) And (lY < m_lHeight - 1 And lY >= 0) then
			m_objImage.Position = (CLng(lY) * m_lWidth) + CLng(lX)
			Pixel = Asc(m_objImage.ReadText(1))
		else
			Pixel = 0
		end if
	End Property

	' Internal fast pixel read and write, no bounds checking!
	Private Property Let FastPixel(ByVal lX,ByVal lY,ByVal lValue)
		m_objImage.Position = (CLng(lY) * m_lWidth) + CLng(lX)
		m_objImage.WriteText Chr(lValue)
	End Property

	Private Property Get FastPixel(ByVal lX,ByVal lY)
		m_objImage.Position = (CLng(lY) * m_lWidth) + CLng(lX)
		FastPixel = Asc(m_objImage.ReadText(1))
	End Property

	Public Property Get Width()
		Width = m_lWidth
	End Property
	
	Public Property Get Height()
		Height = m_lHeight
	End Property

	' Replace() - REALLY REALLY SLOW!
	Public Sub Replace(ByVal lOldColour,ByVal lNewColour)
		Dim sOld, sNew
		
		m_objImage.Position = 0

		sOld = Chr(lOldColour)
		sNew = Chr(lNewColour)
		
		while not m_objImage.EOS
			if m_objImage.ReadText(1) = sOld then
				m_objImage.Position = m_objImage.Position - 1
				m_objImage.WriteText sNew
			end if
		wend
	End Sub
	
	' Duplicate part of the image. Much, much faster...
	Public Sub Copy(ByVal lX1,ByVal lY1,ByVal lX2,ByVal lY2,ByVal lX3,ByVal lY3)
		Dim bXY1, bXY2, bXY3, objCopyImage, lCopyHeight, lCopyWidth, lTemp
		
		bXY1 = (lX1 >= 0 And lX1 < m_lWidth) And (lY1 >=0 And lY1 < m_lHeight)
		bXY2 = (lX2 >= 0 And lX2 < m_lWidth) And (lY2 >=0 And lY2 < m_lHeight)
		bXY3 = (lX3 >= 0 And lX3 < m_lWidth) And (lY3 >=0 And lY3 < m_lHeight)
		
		if bXY1 And bXY2 And bXY3 Then
			' Copy the image first
			Set objCopyImage = Server.CreateObject("ADODB.Stream")
			objCopyImage.Type = 2
			objCopyImage.Charset = "x-ansi"
			objCopyImage.Open
			
			m_objImage.Position = 0
			
			m_objImage.CopyTo objCopyImage
			
			' Now copy stuff across
			lCopyWidth = lX2 - lX1
			lCopyHeight = lY2 - lY1
			
			if (lX3 + lCopyWidth) > m_lWidth then lCopyWidth = m_lWidth - lX3
			if (lY3 + lCopyHeight) > m_lHeight then lCopyHeight = m_lHeight - lY3
			
			For lTemp = lY1 To (lY1 + lCopyHeight) - 1
				objCopyImage.Position = (lTemp * m_lWidth) + lX1
				m_objImage.Position = ((lX3 + (lTemp - lY1)) * m_lWidth) + lX3
				objCopyImage.CopyTo m_objImage,lCopyWidth
			Next
		end if
	End Sub

	' Bounds checked flood fill
	Public Sub Flood(ByVal lX,ByVal lY)
		' Grab the pixel value at lX,lY, we look for all pixels this colour
		Dim xStack, yStack, lPointer, lOldColour, lNewValue, lSouth, lWest, lNorth, lMax
		
		lMax = m_lWidth * m_lHeight

		Redim xStack(lMax) ' Maximum number of values possible ever
		Redim yStack(lMax)
		
		lPointer = 1
		
		m_objImage.Position = (lY * m_lHeight) + lX
		xStack(lPointer) = lX
		yStack(lPointer) = lY

		' Work out some stuff to try and make this quicker
		lOldColour = m_objImage.ReadText(1)
		lNewValue = Chr(ForegroundColourIndex)
		
		on error resume next
		
		while lPointer > 0
			' Read the top of the stack
			lX = xStack(lPointer)
			lY = yStack(lPointer)
			m_objImage.Position = (lY * m_lHeight) + lX
			m_objImage.WriteText lNewValue
			
			m_objImage.Position = (lY * m_lHeight) + lX + 1
			if m_objImage.ReadText(1) = lOldColour and (lX + 1) < m_lWidth then
				lPointer = lPointer + 1
				xStack(lPointer) = lX + 1
				yStack(lPointer) = lY
			else
				m_objImage.Position = ((lY + 1) * m_lHeight) + lX
				if m_objImage.ReadText(1) = lOldColour and (lY + 1) < m_lHeight then
					lPointer = lPointer + 1
					xStack(lPointer) = lX
					yStack(lPointer) = lY + 1
				else
					m_objImage.Position = (lY * m_lHeight) + lX - 1
					if m_objImage.ReadText(1) = lOldColour and (lY - 1) >= 0 then
						lPointer = lPointer + 1
						xStack(lPointer) = lX - 1
						yStack(lPointer) = lY
					else
						m_objImage.Position = ((lY - 1) * m_lHeight) + lX
						if m_objImage.ReadText(1) = lOldColour and (lX - 1) >= 0 then
							lPointer = lPointer + 1
							xStack(lPointer) = lX
							yStack(lPointer) = lY - 1
						else
							lPointer = lPointer - 1
						end if
					end if
				end if
			end if
		wend
	End Sub

	' UnboundFlood
	' A little better, basically a flood fill with no bounds checking. Don't use this unless
	' you are absolutely sure that the shape you're filling in is complete with no holes!
	Public Sub UnboundFlood(ByVal lX,ByVal lY)
		' Grab the pixel value at lX,lY, we look for all pixels this colour
		Dim aStack, lPointer, lOldColour, lNewValue, lSouth, lWest, lNorth, lMax
		
		lMax = m_lWidth * m_lHeight

		Redim aStack(lMax) ' Maximum number of values possible ever
		
		lPointer = 1 ' First cab off the rank
		
		m_objImage.Position = (lY * m_lHeight) + lX		
		aStack(lPointer) = m_objImage.Position

		' Work out some stuff to try and make this quicker
		lOldColour = m_objImage.ReadText(1)
		lNewValue = Chr(ForegroundColourIndex)
		
		lSouth = m_lWidth - 2
		lWest = m_lWidth + 2
		lNorth = m_lWidth
		
		on error resume next
		
		while lPointer > 0
			' Read the top of the stack
			m_objImage.Position = aStack(lPointer)
			m_objImage.WriteText lNewValue
			
			if m_objImage.ReadText(1) = lOldColour then
				lPointer = lPointer + 1
				aStack(lPointer) = m_objImage.Position - 1
			else
				m_objImage.Position = m_objImage.Position + lSouth
				if m_objImage.ReadText(1) = lOldColour then
					lPointer = lPointer + 1
					aStack(lPointer) = m_objImage.Position - 1
				else
					m_objImage.Position = m_objImage.Position - lWest
					if m_objImage.ReadText(1) = lOldColour then
						lPointer = lPointer + 1
						aStack(lPointer) = m_objImage.Position - 1
					else
						m_objImage.Position = m_objImage.Position - lNorth
						if m_objImage.ReadText(1) = lOldColour then
							lPointer = lPointer + 1
							aStack(lPointer) = m_objImage.Position - 1
						else
							lPointer = lPointer - 1
						end if
					end if
				end if
			end if
		wend
	End Sub

	' Draw a regular polygon	
	Public Sub RegularPolygon(lX,lY,lSides,lRadius,lAngle)
		Dim Alpha, lTemp, lX2, lY2
		Dim aX(), aY()
		
		ReDim aX(lSides - 1),aY(lSides - 1)
		
		Alpha = (2 * PI) / lSides
		
		For lTemp = 1 to lSides ' lSides
			aX(lTemp - 1) = CInt(lX + lRadius * Sin(Alpha * lTemp + lAngle))
			aY(lTemp - 1) = CInt(lY + lRadius * Cos(Alpha * lTemp + lAngle))
		Next
		
		Polygon aX,aY,True
	End Sub

	' Draw a regular polygon star, skipping just a single point each time
	Public Sub RegularStar(lX,lY,lSides,lRadius,lAngle)
		Dim Alpha, lTemp, lX2, lY2, lIndex
		Dim aX(), aY()
		
		Alpha = (2 * PI) / lSides
		
		if lSides Mod 2 = 0 then
			ReDim aX(lSides / 2 - 1),aY(lSides / 2 - 1)
			lIndex = 0
			For lTemp = 1 to lSides step 2
				aX(lIndex) = CInt(lX + lRadius * Sin(Alpha * lTemp + lAngle))
				aY(lIndex) = CInt(lY + lRadius * Cos(Alpha * lTemp + lAngle))
				lIndex = lIndex + 1
			Next
			Polygon aX,aY,True
			lIndex = 0
			For lTemp = 2 to lSides step 2
				aX(lIndex) = CInt(lX + lRadius * Sin(Alpha * lTemp + lAngle))
				aY(lIndex) = CInt(lY + lRadius * Cos(Alpha * lTemp + lAngle))
				lIndex = lIndex + 1
			next
			Polygon aX,aY,True
		else
			ReDim aX(lSides - 1),aY(lSides - 1)
			For lTemp = 1 to lSides ' lSides
				aX(lTemp - 1) = CInt(lX + lRadius * Sin(Alpha * lTemp * 2 + lAngle))
				aY(lTemp - 1) = CInt(lY + lRadius * Cos(Alpha * lTemp * 2 + lAngle))
			Next
			Polygon aX,aY,True
		end if		
	End Sub
	
	' Copied from 1.0.6
	Public Sub Polygon(aX,aY,bJoin)
		Dim iTemp

		if UBound(aX) <> UBound(aY) then exit sub
		if UBound(aX) < 1 then exit sub ' Must be more than one point
		
		' Draw a series of lines from arrays aX and aY
		for iTemp = 1 to UBound(aX)
			Line aX(iTemp - 1),aY(iTemp - 1),aX(iTemp),aY(iTemp)
		next
		
		if bJoin then
			Line aX(UBound(aX)),aY(UBound(aY)),aX(0),aY(0)
		end if
	End Sub

	' Easy as, err, rectangle?
	' Copied from 1.0.6
	Public Sub PieSlice(lX,lY,lRadius,sinStartAngle,sinArcAngle,bFilled)
		Dim sinActualAngle, sinMidAngle, lX2, lY2, iTemp
		
		Arc lX,lY,lRadius,lRadius,sinStartAngle - 1,sinArcAngle + 1
		AngleLine lX,lY,lRadius,sinStartAngle
		sinActualAngle = sinStartAngle + sinArcAngle
		if sinActualAngle > 360 then
			sinActualAngle = sinActualAngle - 360
		end if
		AngleLine lX,lY,lRadius,sinActualAngle
		' Now pick a start flood point at the furthest point from the center
		' Divide the arc angle by 2
		sinMidAngle = sinStartAngle + (sinArcAngle / 2)
		
		if sinMidAngle > 360 then
			sinMidAngle = sinMidAngle - 360
		end if

		if bFilled then
			iTemp = lRadius -2
			lY2 = CInt(lY + (Sin(DegreesToRadians(sinMidAngle)) * iTemp))
			lX2 = CInt(lX + (Cos(DegreesToRadians(sinMidAngle)) * iTemp))

			UnboundFlood lX2,lY2
			AngleLine lX,lY,lRadius,sinMidAngle
		end if
	End Sub

	' Copied from 1.0.6
	Public Sub Bezier(lX1,lY1,lCX1,lCY1,lCX2,lCY2,lX2,lY2,lPointCount)
		Dim sinT, lX, lY, lLastX, lLastY, sinResolution
		
		if lPointCount = 0 then exit sub
		
		sinResolution = 1 / lPointCount
		
		sinT = 0
		
		lLastX = lX1
		lLastY = lY1
		
		while sinT <= 1
			lX = int(((sinT^3) * -1 + (sinT^2) * 3 + sinT * -3 + 1) * lX1 + ((sinT^3) *  3 + (sinT^2) *-6 + sinT *  3) * lCX1 + ((sinT^3) * -3 + (sinT^2) * 3) * lCX2 + (sinT^3) * lX2)
			lY = int(((sinT^3) * -1 + (sinT^2) * 3 + sinT * -3 + 1) * lY1 + ((sinT^3) *  3 + (sinT^2) *-6 + sinT *  3) * lCY1 + ((sinT^3) * -3 + (sinT^2) * 3) * lCY2 + (sinT^3) * lY2)

			Line lLastX,lLastY,lX,lY
			
			lLastX = lX
			lLastY = lY
			
			sinT = sinT + sinResolution
		wend

		Line lLastX,lLastY,lX2,lY2
	End Sub

	' ArcPixel Kindly donated by Richard Deeming (www.trinet.co.uk)
	' Copied from 1.0.6
	Private Sub ArcPixel(lX, lY, ltX, ltY, sinStart, sinEnd)
		Dim dAngle
	    
	    If ltX = 0 Then
	        dAngle = Sgn(ltY) * PI / 2
	    ElseIf ltX < 0 And ltY < 0 Then
	        dAngle = PI + Atn(ltY / ltX)
	    ElseIf ltX < 0 Then
	        dAngle = PI - Atn(-ltY / ltX)
	    ElseIf ltY < 0 Then
	        dAngle = 2 * PI - Atn(-ltY / ltX)
	    Else
	        dAngle = Atn(ltY / ltX)
	    End If
	    
	    If dAngle < 0 Then dAngle = 2 * PI + dAngle

		' Compensation for radii spanning over 0 degree marker
		if sinEnd > DegreesToRadians(360) and dAngle < (sinEnd - DegreesToRadians(360)) then
			dAngle = dAngle + DegreesToRadians(360)
		end if
		
	    If sinStart < sinEnd And (dAngle > sinStart And dAngle < sinEnd) Then
	        'This is the "corrected" angle
	        'To change back, change the minus to a plus
	        Pixel(lX + ltX, lY + ltY) = ForegroundColourIndex
	    End If
	End Sub

	' Arc Kindly donated by Richard Deeming (www.trinet.co.uk), vast improvement on the
	' previously kludgy Arc function.
	' Copied from 1.0.6
	Public Sub Arc(ByVal lX, ByVal lY, ByVal lRadiusX, ByVal lRadiusY, ByVal sinStartAngle, ByVal sinArcAngle)
		' Draw an arc at point lX,lY with radius lRadius
		' running from sinStartAngle degrees for sinArcAngle degrees
		Dim lAlpha, lBeta, S, T, lTempX, lTempY, dStart, dEnd
	    
	    dStart = DegreesToRadians(sinStartAngle)
	    dEnd = dStart + DegreesToRadians(sinArcAngle)
	    
	    lAlpha = lRadiusX * lRadiusX
	    lBeta = lRadiusY * lRadiusY
	    lTempX = 0
	    lTempY = lRadiusY
	    S = lAlpha * (1 - 2 * lRadiusY) + 2 * lBeta
	    T = lBeta - 2 * lAlpha * (2 * lRadiusY - 1)
	    ArcPixel lX, lY, lTempX, lTempY, dStart, dEnd
	    ArcPixel lX, lY, -lTempX, lTempY, dStart, dEnd
	    ArcPixel lX, lY, lTempX, -lTempY, dStart, dEnd
	    ArcPixel lX, lY, -lTempX, -lTempY, dStart, dEnd

	    Do
	        If S < 0 Then
	            S = S + 2 * lBeta * (2 * lTempX + 3)
	            T = T + 4 * lBeta * (lTempX + 1)
	            lTempX = lTempX + 1
	        ElseIf T < 0 Then
	            S = S + 2 * lBeta * (2 * lTempX + 3) - 4 * lAlpha * (lTempY - 1)
	            T = T + 4 * lBeta * (lTempX + 1) - 2 * lAlpha * (2 * lTempY - 3)
	            lTempX = lTempX + 1
	            lTempY = lTempY - 1
	        Else
	            S = S - 4 * lAlpha * (lTempY - 1)
	            T = T - 2 * lAlpha * (2 * lTempY - 3)
	            lTempY = lTempY - 1
	        End If

	        ArcPixel lX, lY, lTempX, lTempY, dStart, dEnd
	        ArcPixel lX, lY, -lTempX, lTempY, dStart, dEnd
	        ArcPixel lX, lY, lTempX, -lTempY, dStart, dEnd
	        ArcPixel lX, lY, -lTempX, -lTempY, dStart, dEnd

	    Loop While lTempY > 0
	End Sub

	' Copied from 1.0.6
	Public Sub AngleLine(ByVal lX,ByVal lY,ByVal lRadius,ByVal sinAngle)
		' Draw a line at an angle
		' Angles start from the top vertical and work clockwise
		' Work out the destination defined by length and angle
		Dim lX2, lY2
		
		lY2 = (Sin(DegreesToRadians(sinAngle)) * lRadius)
		lX2 = (Cos(DegreesToRadians(sinAngle)) * lRadius)
		
		Line lX,lY,lX + lX2,lY + lY2
	End Sub

	' Bresenham line algorithm, this is pretty quick, only uses point to point to avoid the
	' mid-point problem. Copied from 1.0.6
	Public Sub Line(ByVal lX1,ByVal lY1,ByVal lX2,ByVal lY2)
		Dim lDX, lDY, lXIncrement, lYIncrement, lDPr, lDPru, lP
		
		lDX = Abs(lX2 - lX1)
		lDY = Abs(lY2 - lY1)
		
		if lX1 > lX2 then
			lXIncrement = -1
		else
			lXIncrement = 1
		end if
		
		if lY1 > lY2 then
			lYIncrement = -1
		else
			lYIncrement = 1
		end if
		
		if lDX >= lDY then
			lDPr = ShiftLeft(lDY,1)
			lDPru = lDPr - ShiftLeft(lDX,1)
			lP = lDPr - lDX
			
			while lDX >= 0
				Pixel(lX1,lY1) = ForegroundColourIndex
				if lP > 0 then
					lX1 = lX1 + lXIncrement
					lY1 = lY1 + lYIncrement
					lP = lP + lDPru
				else
					lX1 = lX1 + lXIncrement
					lP = lP + lDPr
				end if
				lDX = lDX - 1
			wend
		else
			lDPr = ShiftLeft(lDX,1)
			lDPru = lDPr - ShiftLeft(lDY,1)
			lP = lDPR - lDY
			
			while lDY >= 0
				Pixel(lX1,lY1) = ForegroundColourIndex
				if lP > 0 then
					lX1 = lX1 + lXIncrement
					lY1 = lY1 + lYIncrement
					lP = lP + lDPru
				else
					lY1 = lY1 + lYIncrement
					lP = lP + lDPr
				end if
				lDY = lDY - 1
			wend
		end if
	End Sub

	' Draw a custom pattern line fed by sPattern
	Public Sub CustomLine(ByVal lX1,ByVal lY1,ByVal lX2,ByVal lY2,sPattern)
		Dim lDX, lDY, lXIncrement, lYIncrement, lDPr, lDPru, lP, lPatternIndex
		
		if sPattern = "" then sPattern = "1"
		
		lDX = Abs(lX2 - lX1)
		lDY = Abs(lY2 - lY1)
		
		if lX1 > lX2 then
			lXIncrement = -1
		else
			lXIncrement = 1
		end if
		
		if lY1 > lY2 then
			lYIncrement = -1
		else
			lYIncrement = 1
		end if
		
		lPatternIndex = 1

		if lDX >= lDY then
			lDPr = ShiftLeft(lDY,1)
			lDPru = lDPr - ShiftLeft(lDX,1)
			lP = lDPr - lDX
			
			while lDX >= 0
				Pixel(lX1,lY1) = CLng(Mid(sPattern,lPatternIndex,1))
				lPatternIndex = lPatternIndex + 1
				if lPatternIndex > Len(sPattern) then lPatternIndex = 1
				if lP > 0 then
					lX1 = lX1 + lXIncrement
					lY1 = lY1 + lYIncrement
					lP = lP + lDPru
				else
					lX1 = lX1 + lXIncrement
					lP = lP + lDPr
				end if
				lDX = lDX - 1
			wend
		else
			lDPr = ShiftLeft(lDX,1)
			lDPru = lDPr - ShiftLeft(lDY,1)
			lP = lDPR - lDY
			
			while lDY >= 0
				Pixel(lX1,lY1) = CLng(Mid(sPattern,lPatternIndex,1))
				lPatternIndex = lPatternIndex + 1
				if lPatternIndex > Len(sPattern) then lPatternIndex = 1
				if lP > 0 then
					lX1 = lX1 + lXIncrement
					lY1 = lY1 + lYIncrement
					lP = lP + lDPru
				else
					lY1 = lY1 + lYIncrement
					lP = lP + lDPr
				end if
				lDY = lDY - 1
			wend
		end if
	End Sub

	' Copied from 1.0.6
	Public Sub Rectangle(ByVal lX1,ByVal lY1,ByVal lX2,ByVal lY2)
		' Easy as pie, well, actually pie is another function... draw four lines
		Line lX1,lY1,lX2,lY1
		Line lX2,lY1,lX2,lY2
		Line lX2,lY2,lX1,lY2
		Line lX1,lY2,lX1,lY1
	End Sub

	Public Sub FilledRectangle(ByVal lX1,ByVal lY1,ByVal lX2,ByVal lY2)
		' Make a filled rectangle really quickly
		Dim lWidth, lHeight, lTemp
		
		lWidth = lX2 - lX1
		lHeight = lY2 - lY1
		
		if lX2 > m_lWidth then lWidth = m_lWidth - lX1
		if lY2 > m_lHeight then lHeight = m_lHeight - lY1
		
		For lTemp = lY1 to (lY1 + lHeight)
			m_objImage.Position = (lTemp * m_lWidth) + lX1
			m_objImage.WriteText String(lWidth,Chr(ForegroundColourIndex))
		Next
	End Sub

	' Copied from 1.0.6
	Public Sub Circle(ByVal lX,ByVal lY,ByVal lRadius)
		Ellipse lX,lY,lRadius,lRadius
	End Sub

	' Bresenham ellispe, pretty quick also, uses reflection, so rotation is out of the 
	' question unless we perform a matrix rotation after rendering the ellipse coords
	' copied from 1.0.6
	Public Sub Ellipse(ByVal lX,ByVal lY,ByVal lRadiusX,ByVal lRadiusY)
		' Draw a circle at point lX,lY with radius lRadius
		Dim lAlpha,lBeta,S,T,lTempX,lTempY
		
		lAlpha = lRadiusX * lRadiusX
		lBeta = lRadiusY * lRadiusY
		lTempX = 0
		lTempY = lRadiusY
		S = lAlpha * (1 - 2 * lRadiusY) + 2 * lBeta
		T = lBeta - 2 * lAlpha * (2 * lRadiusY - 1)
		Pixel(lX + lTempX,lY + lTempY) = ForegroundColourIndex
		Pixel(lX - lTempX,lY + lTempY) = ForegroundColourIndex
		Pixel(lX + lTempX,lY - lTempY) = ForegroundColourIndex
		Pixel(lX - lTempX,lY - lTempY) = ForegroundColourIndex
		Do
			if S < 0 then
				S = S + 2 * lBeta * (2 * lTempX + 3)
				T = T + 4 * lBeta * (lTempX + 1)
				lTempX = lTempX + 1
			elseif T < 0 then
				S = S + 2 * lBeta * (2 * lTempX + 3) - 4 * lAlpha * (lTempY - 1)
				T = T + 4 * lBeta * (lTempX + 1) - 2 * lAlpha * (2 * lTempY - 3)
				lTempX = lTempX + 1
				lTempY = lTempY - 1
			else
				S = S - 4 * lAlpha * (lTempY - 1)
				T = T - 2 * lAlpha * (2 * lTempY - 3)
				lTempY = lTempY - 1
			end if
			Pixel(lX + lTempX,lY + lTempY) = ForegroundColourIndex
			Pixel(lX - lTempX,lY + lTempY) = ForegroundColourIndex
			Pixel(lX + lTempX,lY - lTempY) = ForegroundColourIndex
			Pixel(lX - lTempX,lY - lTempY) = ForegroundColourIndex
		loop while lTempY > 0
	End Sub

	' Vector font support
	' These fonts are described in terms of points on a grid with simple
	' X and Y offsets. These functions take elements of a string and render
	' them from arrays storing character vector information. Vector fonts are
	' have proportional widths, unlike bitmapped fonts which are fixed in size
	' The format for the vector array is simply a variable length list of x y pairs
	' the sub DrawVectorChar renders the single character from the array.
	' The other advantage of vector fonts is that they can be scaled :)

	' Maybe add an angle value?
	Public Sub DrawVectorTextWE(ByVal lX,ByVal lY,sText,lSize)
		Dim iTemp, lCurrentStringX
		
		lCurrentStringX = lX
		
		For iTemp = 1 to Len(sText)
			lCurrentStringX = lCurrentStringX + DrawVectorChar(lCurrentStringX,lY,Mid(sText,iTemp,1),lSize,true) + int(lSize)
		Next
	End Sub
	
	Public Sub DrawVectorTextNS(ByVal lX,ByVal lY,sText,lSize)
		Dim iTemp, lCurrentStringY
		
		lCurrentStringY = lY
		
		For iTemp = 1 to Len(sText)
			lCurrentStringY = lCurrentStringY + DrawVectorChar(lX,lCurrentStringY,Mid(sText,iTemp,1),lSize,false) + int(lSize)
		Next
	End Sub
	
	Private Function DrawVectorChar(ByVal lX,ByVal lY,sChar,lSize,bOrientation)
		Dim iTemp, aFont, lLargest
		
		if IsArray(VLetter) then
			if sChar <> " " then
				aFont = VFont(sChar)
		
				if bOrientation then
					lLargest = aFont(1,0) * lSize
				else
					lLargest = aFont(1,1) * lSize
				end if
		
				for iTemp = 1 to UBound(aFont,1) - 1
					if bOrientation then
						if aFont(iTemp,2) = 1  then ' Pen down
							Line lX + aFont(iTemp - 1,0) * lSize,lY + aFont(iTemp - 1,1) * lSize,lX + aFont(iTemp,0) * lSize,lY + aFont(iTemp,1) * lSize
						end if
						if (aFont(iTemp,0) * lSize) > lLargest then
							lLargest = aFont(iTemp,0) * lSize
						end if
					else
						if aFont(iTemp,2) = 1 then ' Pen down
							Line lX + aFont(iTemp - 1,0) * lSize,lY + aFont(iTemp - 1,1) * lSize,lX + aFont(iTemp,0) * lSize,lY + aFont(iTemp,1) * lSize
						end if
						if (aFont(iTemp,1) * lSize) > lLargest then
							lLargest = aFont(iTemp,1) * lSize
						end if
					end if
				next
			else
				lLargest = lSize * 3
			end if
		else
			Alert "font.asp has not been included into this canvas project, this is needed for displaying fonts",false
		end if		
		' Return the width of the character
		DrawVectorChar = lLargest
	End Function

	' Return the pixel size of the text going WE and NS
	Public Function GetTextWEWidth(sText)
		Dim lWidth, lTemp
		' Tricky, get the width of all the characters
		lWidth = 0
		For lTemp = 1 to Len(sText)
			lWidth = lWidth + Len(Font(Mid(sText,lTemp,1))(0))
		Next
		
		GetTextWEWidth = lWidth
	End Function
	
	' All characters have the same height in the fonts
	Public Function GetTextWEHeight(sText)
		GetTextWEHeight = UBound(Letter) + 1
	End Function
	
	' Find the widest character
	Public Function GetTextNSWidth(sText)
		Dim lWidth, lTemp
		lWidth = 0
		For lTemp = 1 to Len(sText)
			if Len(Font(Mid(sText,lTemp,1))(0)) > lWidth then
				lWidth = Len(Font(Mid(sText,lTemp,1))(0))
			end if
		Next
		
		GetTextNSWidth = lWidth
	End Function
	
	Public Function GetTextNSHeight(sText)
		' Easy one, multiply the number of character by the Letter array size
		GetTextNSHeight = Len(sText) * (UBound(Letter) + 1)
	End Function

	' Bitmap font support
	Public Sub DrawTextWE(ByVal lX,ByVal lY,sText)
		' Render text at lX,lY
		' There's a global dictionary object called Font and it should contain all the 
		' letters in arrays of a 5x5 grid
		Dim iTemp1, iTemp2, iTemp3, bChar, lCurrentX
		
		if IsArray(Letter) then
			For iTemp1 = 0 to UBound(Letter) - 1
				lCurrentX = lX
				For iTemp2 = 1 to len(sText)
					For iTemp3 = 1 to Len(Font(Mid(sText,iTemp2,1))(iTemp1))
						bChar = Mid(Font(Mid(sText,iTemp2,1))(iTemp1),iTemp3,1)
						if bChar <> "0" then Pixel(lCurrentX,lY + iTemp1) = CLng(bChar)
						lCurrentX = lCurrentX + 1
					next
				next
			next
		else
			Alert "font.asp has not been included into this canvas project, this is needed for displaying fonts",false
		end if
	End Sub

	Public Sub DrawTextNS(ByVal lX,ByVal lY,sText)
		' Render text at lX,lY
		' There's a global dictionary object called Font and it should contain all the 
		' letters in arrays of a 5x5 grid
		Dim iTemp1, iTemp2, iTemp3, bChar

		if IsArray(Letter) then
			for iTemp1 = 1 to len(sText)
				for iTemp2 = 0 to UBound(Letter) - 1
					for iTemp3 = 1 to len(Font(Mid(sText,iTemp1,1))(iTemp2))
						bChar = Mid(Font(Mid(sText,iTemp1,1))(iTemp2),iTemp3,1)
						if bChar <> "0" then
							Pixel(lX + iTemp3,lY + ((iTemp1 - 1) * (UBound(Letter) + 1)) + iTemp2) = CLng(bChar)
						end if
					next
				next
			next
		else
			Alert "font.asp has not been included into this canvas project, this is needed for displaying fonts",false
		end if
	End Sub

	Public Sub Clear()
		' Clear the image
		m_objImage.Close
		m_objImage.Open
		m_objImage.WriteText String(m_lWidth * m_lHeight,Chr(BackgroundColourIndex))
	End Sub

	public sub Resize(ByVal lNewWidth,ByVal lNewHeight,bPreserve)
		Dim objNewImage, lCopyWidth, lCopyHeight, lTemp
		
		if bPreserve then
			' Make the new canvas
			Set objNewImage = Server.CreateObject("ADODB.Stream")
			objNewImage.Type = 2
			objNewImage.Charset = "x-ansi"
			objNewImage.Open
		
			objNewImage.WriteText String(lNewWidth * lNewHeight,Chr(BackgroundColourIndex))
			
			lCopyWidth = m_lWidth
			lCopyHeight = m_lHeight
			
			if m_lWidth >= lNewWidth then lCopyWidth = lNewWidth			
			if m_lHeight >= lNewHeight then lCopyHeight = lNewHeight
			
			For lTemp = 0 To lCopyHeight
				m_objImage.Position = (lTemp * m_lWidth)
				objNewImage.Position = (lTemp * lNewWidth)
				m_objImage.CopyTo objNewImage,lCopyWidth
			Next
			
			m_objImage.Close
			m_objImage.Open
			objNewImage.Position = 0
			objNewImage.CopyTo m_objImage
			m_lWidth = lNewWidth
			m_lHeight = lNewHeight
		else
			m_lWidth = lNewWidth
			m_lHeight = lNewHeight
			
			Clear
		end if
	end sub

	Private Sub Alert(sText,bAllowImage)
		' Dump errors out
		if ErrorsToImage and bAllowImage then
			GlobalColourTable(0) = RGB(255,255,255)
			GlobalColourTable(1) = RGB(255,128,0)
			Resize 800,30,False
			ForegroundColourIndex = 1
			DrawTextWE 0,0,"ERROR: " & sText
			DrawTextWE 1,0,"ERROR: " & sText
			Write
		end if
		if ErrorsToResponse and not errorstoimage then Response.Write sText & "<br>"
		if ErrorsToIISLog then Response.AppendToLog sText
		if ErrorsHalt then Response.End
	End Sub

	' 0-0-0-0-0-0-0-0-0-0-0-0-0-0- Image management functions -0-0-0-0-0-0-0-0-0-0-0-0-0-0
	
	' Dump the image
	Public Sub Write()
		Response.ContentType = "image/gif"
		Response.AddHeader "Content-Disposition","filename=SecurityImage.gif"
		Response.BinaryWrite ImageData().Read
	End Sub
	
	Public Sub SaveGIF(sFilename)
		ImageData.SaveToFile sFilename,1 ' adSaveCreateNotExist
	End Sub
	
	' Raw image data
	Public Function ImageData()
		Dim objTemp, lTemp, sText, objTempStream, lSize
		
		Set objTemp = Server.CreateObject("ADODB.Stream")
		
		' Do this all in x-ansi and then convert to binary
		objTemp.Type = 2
		objTemp.Charset = "x-ansi" ' x-ansi is the solution!
		objTemp.Open
		
		objTemp.WriteText "GIF89a"
		
		objTemp.WriteText MakeWord(m_lWidth)
		objTemp.WriteText MakeWord(m_lHeight)
		objTemp.WriteText MakeByte(&HF7)
		objTemp.WriteText MakeByte(BackgroundColourIndex)
		objTemp.WriteText MakeByte(&H00)
		
		For lTemp = 0 To UBound(GlobalColourTable) - 1
			objTemp.WriteText MakeByte(Red(GlobalColourTable(lTemp)))
			objTemp.WriteText MakeByte(Green(GlobalColourTable(lTemp)))
			objTemp.WriteText MakeByte(Blue(GlobalColourTable(lTemp)))
		Next
		
		if UseTransparency then
			objTemp.WriteText MakeWord(&HF921)
			objTemp.WriteText MakeWord(&H0104)
			objTemp.WriteText MakeWord(&H0000)
			objTemp.WriteText MakeByte(TransparentColourIndex)
			objTemp.WriteText MakeByte(&H00)
		end if
		
		if Comment <> "" then
			objTemp.WriteText MakeWord(&HFE21)
			sText = Left(Comment,&HFF) ' Truncate to 255 characters
			objTemp.WriteText MakeByte(Len(sText))
			objTemp.WriteText sText
			objTemp.WriteText MakeByte(&H00)
		end if

		objTemp.WriteText ","
		objTemp.WriteText MakeWord(&H00)
		objTemp.WriteText MakeWord(&H00)
		objTemp.WriteText MakeWord(m_lWidth)
		objTemp.WriteText MakeWord(m_lHeight)
		objTemp.WriteText MakeWord(&H0700)		

		Set objTempStream = Server.CreateObject("ADODB.Stream")
		
		objTempStream.Charset = "x-ansi"
		objTempStream.Type = 2
		objTempStream.Open
		
		m_objImage.Position = 0
		
		' Insert clearcodes
		while not m_objImage.EOS
			objTempStream.WriteText m_objImage.ReadText(&H7E)
			objTempStream.WriteText Chr(&H80)
		wend
		
		objTempStream.Position = 0

		while not objTempStream.EOS
			lSize = objTempStream.Size - objTempStream.Position
			if lSize > &HFF then lSize = &HFF
			objTemp.WriteText MakeByte(lSize)
			objTempStream.CopyTo objTemp,lSize
		wend
		
		objTemp.WriteText MakeWord(&H8100)
		objTemp.WriteText MakeWord(&H3B00)
		
		objTemp.Position = 0
		objTemp.Type = 1
		
		Set ImageData = objTemp
	End Function

	Public Sub LoadBMP(sFilename)
		Dim objStream, lOffset, lNewWidth, lNewHeight, lBPP, lCompression, lImageSize, lTemp
		Dim lRed,lGreen,lBlue,lPad, lLineSize
		
		Set objStream = Server.CreateObject("ADODB.Stream")
		
		objStream.Charset = "x-ansi"
		objStream.Type = 2
		objStream.Open
		
		on error resume next
		
		objStream.LoadFromFile sFilename
		
		if err.number <> 0 then
			alert "Failure calling LoadBMP: " & err.description & "-" & sFilename,true
		end if
		
		on error goto 0
		
		LoadBMPFromStream objStream

		objStream.Close

		Set objStream = Nothing		
	End Sub

	Public Sub LoadBMPFromStream(objStream)
		Dim lOffset, lNewWidth, lNewHeight, lBPP, lCompression, lImageSize, lTemp
		Dim lRed,lGreen,lBlue,lPad, lLineSize
		
		' Decode the bitmap here
		if objStream.ReadText(2) = "BM" then
			objStream.Position = 10
			lOffset = GetLong(objStream.ReadText(4))
			objStream.Position = 18
			lNewWidth = GetLong(objStream.ReadText(4))
			lNewHeight = GetLong(objStream.ReadText(4))
			objStream.Position = 28
			lBPP = GetWord(objStream.ReadText(2))
			lCompression = GetLong(objStream.ReadText(4))
			lImageSize = GetLong(objStream.ReadText(4))
			
			if lBPP = 8 And lCompression = 0 then
				Resize lNewWidth,lNewHeight,False
				objStream.Position = 54
				For lTemp = 0 to 255
					lBlue = Asc(objStream.ReadText(1))
					lGreen = Asc(objStream.ReadText(1))
					lRed = Asc(objStream.ReadText(1))
					lPad = Asc(objStream.ReadText(1))
					GlobalColourTable(lTemp) = RGB(lRed,lGreen,lBlue)
					if lTemp > 127 and GlobalColourTable(lTemp) > 0 and BMPWarnings then
						alert "This image contains more than 128 valid colour entries",true
					end if
				Next
				lPad = 4 - (lNewWidth Mod 4)
				
				if lPad = 4 then lPad = 0
				
				lLineSize = lNewWidth + lPad
				
				' Copy the bitmap image in backwards
				objStream.Position = lOffset 
				For lTemp = lNewHeight - 1 to 0 step - 1
					m_objImage.Position = (lTemp * lNewWidth)
					objStream.CopyTo m_objImage, lNewWidth
					objStream.ReadText lPad
				Next
			elseif BMPWarnings then
				if lBPP <> 8 then alert "This image is not an 8 bit (256/128 colour) image",true
				if lCompression <> 0 then alert "This image is compressed, ASPCanvas cannot load",true
			end if
		elseif BMPWarnings then
			alert "This is not a valid BMP file",true
		end if
	End Sub

	' Force a rough reduction in colours from 256 to 128
	' the mapping occurs from copying the pixels from 
	' one stream to the other with different charsets
	Public Sub TruncatePalette()
		Dim lTemp, sTemp, objTemp

		' Reset all colours past index 127
		For lTemp = 128 to 255
			GlobalColourTable(lTemp) = 0
		Next

		' At this point we cheat a little, we actually create
		' a stream with the ascii charset, this will cause
		' any pixels over the value of 127 to be mapped to various
		' values below 127. It makes this whole thing really quick.
		Set objTemp = Server.CreateObject("ADODB.Stream")
		objTemp.Charset = "ascii" ' Only the first 127 characters are un-filtered
		objTemp.Type = 2
		objTemp.Open
		
		m_objImage.Position = 0
		m_objImage.CopyTo objTemp

		m_objImage.Position = 0
		objTemp.Position = 0
		objTemp.CopyTo m_objImage
	End Sub

	Public Sub WriteBMP()
		Dim objStream
		
		Response.ContentType = "image/bmp"
		Response.AddHeader "Content-Disposition","filename=image.bmp"
		
		Set objStream = MakeBMPStream()
		
		objStream.Position = 0
		objStream.Type = 1
		
		Response.BinaryWrite objStream.Read()
	End Sub

	Public Sub SaveBMP(sFilename)
		Dim objStream
		
		Set objStream = MakeBMPStream()
		
		on error resume next
		
		objStream.SaveToFile sFilename,2

		if err.number <> 0 then
			alert "Failure calling SaveBMP: " & err.description & "-" & sFilename,true
		end if

		on error goto 0

		objStream.Close
		
		Set objStream = Nothing
	End Sub
	
	Private Function MakeBMPStream()
		Dim objStream, lPad, lImageSize, lTemp, sPad
		
		Set objStream = Server.CreateObject("ADODB.Stream")
		
		objStream.Charset = "x-ansi"
		objStream.Type = 2
		objStream.Open
		objStream.WriteText "BM"
		objStream.WriteText MakeLong(&H0000)
		objStream.WriteText MakeLong(&H0000)
		objStream.WriteText MakeLong(&H0436)
		objStream.WriteText MakeLong(&H0028)
		objStream.WriteText MakeLong(m_lWidth)
		objStream.WriteText MakeLong(m_lHeight)
		objStream.WriteText MakeWord(&H0001)
		objStream.WriteText MakeWord(&H0008)
		objStream.WriteText MakeLong(&H0000)
		
		lPad = 4 - (m_lWidth Mod 4)
		
		if lPad = 4 then lPad = 0
		
		lImageSize = (m_lWidth + lPad) * m_lHeight
		
		objStream.WriteText MakeLong(lImageSize)
		
		objStream.WriteText MakeLong(&H0000)
		objStream.WriteText MakeLong(&H0000)
		objStream.WriteText MakeLong(&H00FF)
		objStream.WriteText MakeLong(&H00FF)
		
		For lTemp = 0 to UBound(GlobalColourTable) - 1
			objStream.WriteText MakeByte(Blue(GlobalColourTable(lTemp)))
			objStream.WriteText MakeByte(Green(GlobalColourTable(lTemp)))
			objStream.WriteText MakeByte(Red(GlobalColourTable(lTemp)))
			objStream.WriteText MakeByte(0)
		Next
		
		sPad = String(lPad,Chr(0))
		
		For lTemp = m_lHeight to 0 step - 1
			m_objImage.Position = lTemp * m_lWidth
			m_objImage.CopyTo objStream, m_lWidth
			objStream.WriteText sPad
		Next
		
		objStream.Position = 3
		objStream.WriteText MakeLong(objStream.Size)

		' Copy the stream
		Set MakeBMPStream = objStream
	End Function
	
	' Create a font pack from a loaded BMP file and output the text
	' Fonts are defined from Chr(33) to Chr(126)
	' Each character must be seperated by colour index 2 (CI2)
	' Font colour must be in colour index 1 (CI1)
	' Background colour must be in colour index 0 (CI0)
	Public Function CreateFontPack(lSpaceWidth, lBorder)
		Dim lTemp, lChar, lCharStart, lCharEnd, sLine, lPixel, sTemp, lTempY, aFont()
		sTemp = sTemp & "&lt;%" & vbCRLF
		sTemp = sTemp & "' ASPCanvas FontMaker generated font" & vbCRLF
		sTemp = sTemp & "' Bitmapped font only" & vbCRLF
		sTemp = sTemp & "Dim Font,Letter(" & Height & ")" & vbCRLF
		sTemp = sTemp & "Set Font = Server.CreateObject(""Scripting.Dictionary"")" & vbCRLF & vbCRLF

		' Space first
		For lTemp = 0 To Height - 1
			sTemp = sTemp & "Letter(" & lTemp & ") = """ & String(lSpaceWidth,"0") & """" & vbCRLF
		Next

		sTemp = sTemp & "Font.Add "" "",Letter" & vbCRLF

		ReDim aFont(127,Height)

		For lTempY = 0 to Height - 1
			lChar = 33
			bFound = False
			For lTemp = 0 to Width - 1
				lPixel = Pixel(lTemp,lTempY)
				Select Case lPixel
					Case 0,1
						bFound = True
						sLine = sLine & CStr(lPixel)
					Case Else
						if bFound then 
							aFont(lChar,lTempY) = String(lBorder,"0") & sLine & String(lBorder,"0")
							lChar = lChar + 1
							sLine = ""
						end if
						bFound = False
				End Select
			Next
		Next
		
		For lTempY = 33 to 126
			For lTemp = 0 To Height - 1
				sTemp = sTemp & "Letter(" & lTemp & ") = """ & aFont(lTempY,lTemp) & """" & vbCRLF
			Next
			sTemp = sTemp & "Font.Add Chr(" & lTempY & "),Letter" & vbCRLF & vbCRLF
		Next
		
		sTemp = sTemp & "%&gt;" & vbCRLF		

		CreateFontPack = sTemp
	End Function
	
	' Shift all bits left, big improvement on the old shifts. This does it properly
	' without tripping on the sign at the end of the word
	Private Function ShiftLeft(lValue,iBits)
		' Guilty until proven innocent
		ShiftLeft = 0

		If iBits = 0 then
			ShiftLeft = lValue ' No shifting to do
		ElseIf iBits = 31 Then ' Quickly shift left if there is a value, being aware of the sign
			If lValue And 1 Then
				ShiftLeft = &H80000000
			End If
		Else ' Shift left x bits, being careful with the sign
			If (lValue And m_lPow(31 - iBits)) Then
				ShiftLeft = ((lValue And m_lMask(31 - (iBits + 1))) * m_lPow(iBits)) Or &H80000000
			Else
				ShiftLeft = ((lValue And m_lMask(31 - iBits)) * m_lPow(iBits))
			End If
		End If
	End Function

	Private Function ShiftRight(lValue,iBits)
		' Guilty until proven innocent
		ShiftRight = 0
			
		If iBits = 0 then
			ShiftRight = lValue ' No shifting to do
		ElseIf iBits = 31 Then ' Quickly shift to the right if there is a value in the sign
			If lValue And &H80000000 Then
				ShiftRight = 1
			End If
		Else
			ShiftRight = (lValue And &H7FFFFFFE) \ m_lPow(iBits)

			If (lValue And &H80000000) Then
				ShiftRight = (ShiftRight Or (&H40000000 \ m_lPow(iBits - 1)))
			End If
		End If
	End Function

	' Low byte order
	Private function Low(lValue)
		Low = lValue and &HFF&
	end function

	' High byte order
	Private function High(lValue)
		High = ShiftRight(lValue,8) And &HFF&
	end function

	Private function Blue(lValue)
		Blue = Low(ShiftRight(lValue,16))
	end function

	Private function Green(lValue)
		Green = Low(ShiftRight(lValue,8))
	end function

	Private function Red(lValue)
		Red = Low(lValue)
	end function

	Private function GetLong(sValue)
		GetLong = 0
		if LenB(sValue) >= 4 then
			GetLong = ShiftLeft(GetWord(Mid(sValue,3,2)),16) or GetWord(Mid(sValue,1,2))
		end if
	end function

	Private function MakeLong(lValue)
		Dim lLowWord, lHighWord
		
		lLowWord = lValue and 65535
		lHighWord = ShiftRight(lValue,16) and 65535
		
		MakeLong = MakeWord(lLowWord) & MakeWord(lHighWord)
	end function

	' Get a number from a big-endian word
	Private function GetWord(sValue)
		GetWord = ShiftLeft(Asc(Right(sValue,1)),8) or Asc(Left(sValue,1))
	end function

	' Make a big-endian word
	Private function MakeWord(lValue)
		MakeWord = Chr(Low(lValue)) & Chr(High(lValue))
	end function

	' Filter out the high byte
	Private function MakeByte(lValue)
		MakeByte = Chr(Low(lValue))
	end function
	
	Private Function DegreesToRadians(ByVal sinAngle)
		DegreesToRadians = sinAngle * (PI/180)
	End Function

	Private Function RadiansToDegrees(ByVal sinAngle)
		RadiansToDegrees = sinAngle * (180/PI)
	End Function

	public sub WebSafePalette()
		' Reset the colours to the web safe palette
		Dim iTemp1, iTemp2, iTemp3, iIndex
		
		iIndex = 0
		
		For iTemp1 = &HFF0000& to 0 step - &H330000&
			For iTemp2 = &HFF00& to 0 step - &H3300&
				For iTemp3 = &HFF& to 0 step - &H33&
					GlobalColourTable(iIndex) = iTemp1 or iTemp2 or iTemp3
					iIndex = iIndex + 1
				Next
			Next
		Next
	end sub

	Private Sub Class_Initialize()
		Set m_objImage = Server.CreateObject("ADODB.Stream")
		
		m_objImage.Type = 2 ' Text
		m_objImage.Charset = "x-ansi" ' x-ansi, true 8 bit unsigned data
		m_objImage.Open ' Open the stream
		
		m_lWidth = 20
		m_lHeight = 20
		
		ReDim GlobalColourTable(256)
		
		BackgroundColourIndex = 0
		ForegroundColourIndex = 0
		UseTransparency = False
		TransparentColourIndex = 0
		Comment = ""

		ErrorsToIISLog = False
		ErrorsToResponse = True
		ErrorsHalt = True
		ErrorsToImage = True

		BMPWarnings = True

		ReDim m_lMask(30)
		ReDim m_lPow(30)

		m_lMask(0)	=	CLng(&H00000001&)
		m_lMask(1)	=	CLng(&H00000003&)
		m_lMask(2)	=	CLng(&H00000007&)
		m_lMask(3)	=	CLng(&H0000000F&)
		m_lMask(4)	=	CLng(&H0000001F&)
		m_lMask(5)	=	CLng(&H0000003F&)
		m_lMask(6)	=	CLng(&H0000007F&)
		m_lMask(7)	=	CLng(&H000000FF&)
		m_lMask(8)	=	CLng(&H000001FF&)
		m_lMask(9)	=	CLng(&H000003FF&)
		m_lMask(10)	=	CLng(&H000007FF&)
		m_lMask(11)	=	CLng(&H00000FFF&)
		m_lMask(12)	=	CLng(&H00001FFF&)
		m_lMask(13)	=	CLng(&H00003FFF&)
		m_lMask(14)	=	CLng(&H00007FFF&)
		m_lMask(15)	=	CLng(&H0000FFFF&)
		m_lMask(16)	=	CLng(&H0001FFFF&)
		m_lMask(17)	=	CLng(&H0003FFFF&)
		m_lMask(18)	=	CLng(&H0007FFFF&)
		m_lMask(19)	=	CLng(&H000FFFFF&)
		m_lMask(20)	=	CLng(&H001FFFFF&)
		m_lMask(21)	=	CLng(&H003FFFFF&)
		m_lMask(22)	=	CLng(&H007FFFFF&)
		m_lMask(23)	=	CLng(&H00FFFFFF&)
		m_lMask(24)	=	CLng(&H01FFFFFF&)
		m_lMask(25)	=	CLng(&H03FFFFFF&)
		m_lMask(26)	=	CLng(&H07FFFFFF&)
		m_lMask(27)	=	CLng(&H0FFFFFFF&)
		m_lMask(28)	=	CLng(&H1FFFFFFF&)
		m_lMask(29)	=	CLng(&H3FFFFFFF&)
		m_lMask(30)	=	CLng(&H7FFFFFFF&)

		' Use these to speed up multiplication a little where needed
		m_lPow(0)	=	CLng(&H00000001&)
		m_lPow(1)	=	CLng(&H00000002&)
		m_lPow(2)	=	CLng(&H00000004&)
		m_lPow(3)	=	CLng(&H00000008&)
		m_lPow(4)	=	CLng(&H00000010&)
		m_lPow(5)	=	CLng(&H00000020&)
		m_lPow(6)	=	CLng(&H00000040&)
		m_lPow(7)	=	CLng(&H00000080&)
		m_lPow(8)	=	CLng(&H00000100&)
		m_lPow(9)	=	CLng(&H00000200&)
		m_lPow(10)	=	CLng(&H00000400&)
		m_lPow(11)	=	CLng(&H00000800&)
		m_lPow(12)	=	CLng(&H00001000&)
		m_lPow(13)	=	CLng(&H00002000&)
		m_lPow(14)	=	CLng(&H00004000&)
		m_lPow(15)	=	CLng(&H00008000&)
		m_lPow(16)	=	CLng(&H00010000&)
		m_lPow(17)	=	CLng(&H00020000&)
		m_lPow(18)	=	CLng(&H00040000&)
		m_lPow(19)	=	CLng(&H00080000&)
		m_lPow(20)	=	CLng(&H00100000&)
		m_lPow(21)	=	CLng(&H00200000&)
		m_lPow(22)	=	CLng(&H00400000&)
		m_lPow(23)	=	CLng(&H00800000&)
		m_lPow(24)	=	CLng(&H01000000&)
		m_lPow(25)	=	CLng(&H02000000&)
		m_lPow(26)	=	CLng(&H04000000&)
		m_lPow(27)	=	CLng(&H08000000&)
		m_lPow(28)	=	CLng(&H10000000&)
		m_lPow(29)	=	CLng(&H20000000&)
		m_lPow(30)	=	CLng(&H40000000&)
		
		Clear()
	End Sub
	
	Private Sub Class_Terminate()
		Set m_objImage = Nothing
	End Sub
End Class
%>