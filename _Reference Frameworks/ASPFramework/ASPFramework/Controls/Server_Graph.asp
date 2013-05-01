<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::Server Graph

	Public Function New_ServerGraph(name) 
		Set New_ServerGraph = New ServerGraph
			New_ServerGraph.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerGraph"

	Class ServerGraph
		Public Control
		Public Title
		Public Width
		Public Height
		Public Mode  		'1 - Simple Bars, 2 - Stacked Bars, 3 - Stacked Bars (Relative Area)
		Public XScaleTick	'The number of ticks between ticks with labels
		Public XScale
		Public YScale
		Public XLabel
		Public YLabel
		Public YOffset
		Public ShowYear
		Public ShowDay
		Public ImageStyle
		
		Private mRowData()
		Private mLegendData
		Private mLegend
		
		'Time scale
		Private mTScale
		Private mTScaleHour
		Private mTScaleMin
		'Date scale
		Private mDScale
		Private mDScaleDay
		Private mDScaleMon
		Private mDScaleYear
		Private mLongDate
		
		Private Sub Class_Initialize()			
			Set Control = New WebControl	
			Set Control.Owner = Me

			'Initialize variables
			ReDim mRowData(0)
			ReDim mLegendData(0)

			mDScale		= False
			mTScale		= False
			mLegend		= False
			mLongDate	= False

			Mode		= 1
			Width		= 400
			Height		= 200
			Title		= ""
			XScaleTick	= 1
			XScale		= 1
			YScale		= 1
			XLabel		= ""
			YLabel		= ""
			YOffset		= 1
			ShowYear	= False
			ShowDay		= False		
			ImageStyle	= "border:0 black outset;border-bottom:0;"
						
		End Sub
	   		
		Public Function SetValueFromDataSource(value)

		End Function
	   
		Public Function ReadProperties(bag)		
			Mode		= bag.Read("M")
			Width		= bag.Read("W")	
			Height		= bag.Read("H")	
			Title		= bag.Read("TL")	
			XScaleTick	= bag.Read("TK")
			XScale		= bag.Read("XS")
			YScale		= bag.Read("YS")
			XLabel		= bag.Read("XL")	
			YLabel		= bag.Read("YL")
			YOffset		= bag.Read("YO")
			ShowYear	= bag.Read("SY")
			ShowDay		= bag.Read("SD")
			ImageStyle	= bag.Read("IS")
		End Function
		
		Public Function WriteProperties(bag)
			bag.Write "M",Mode	
			bag.Write "W",Width	
			bag.Write "H",Height
			bag.Write "TL",Title		
			bag.Write "TK",XScaleTick	
			bag.Write "XS",XScale	
			bag.Write "YS",YScale	
			bag.Write "XL",XLabel		
			bag.Write "YL",YLabel	
			bag.Write "YO",YOffset	
			bag.Write "SY",ShowYear
			bag.Write "SD",ShowDay	
			bag.Write "IS",ImageStyle			
		End Function
		
		Public Function AddRow(RowValues)		
		Dim arrRow
			'make sure RowValues is an array or numeric Value
			If Not IsArray(RowValues) Then 
				arrRow = Array(RowValues)
			Else
				arrRow = RowValues
			End If
			
			mRowData(UBound(mRowData)) = arrRow
			ReDim Preserve mRowData(UBound(mRowData)+1) 
		End Function

		Public Function SetLegend(LegValues)	
		Dim arrRow
			mLegend = True
			'make sure RowValues is an array or numeric Value
			If Not IsArray(LegValues) Then 
				arrRow = Array(LegValues)
			Else
				arrRow = LegValues
			End If
			
			mLegendData = arrRow
		End Function
		
		Public Function SetTimeScale(iHour,iMin)
			mTScale = True	' activate time scale
			mTScaleHour	= iHour
			mTScaleMin	= iMin
		End Function
		
		Public Function SetDateScale(iMon,iDay,iYear,bLongDate)
			mDScale = True	' activate date scale
			mDScaleDay	= iDay
			mDScaleMon	= iMon
			mDScaleYear	= iYear
			mLongDate	= bLongDate
		End Function
				
		Private Function RenderGraph ()		
		Dim i
		Dim r
			Response.Write "<script language=""javascript"">"  + vbCrLf + _
				"var g = "+Control.Name+"= new Graph("& Width &","& Height &");" + vbCrLf
				' write row data
				For i = 0 To UBound(mRowData)
					r = mRowData(i)
					If IsArray(r) Then
						Response.Write "g.addRow("+Join(r,",")+");" + vbCrLf
					End If
				Next
				'write properties
				Response.Write "g.stacked= "& IIf(Mode=2 Or Mode=3,"true","false")&";" & vbCrLf
				Response.Write "g.relative="& IIf(Mode=3,"true","false")&";" & vbCrLf
				Response.Write "g.inc=" 	& CStr(XScale) &";" & vbCrLf
				Response.Write "g.scale="	& CStr(YScale) &";" & vbCrLf
				Response.Write "g.skip=" 	& CStr(XScaleTick) &";" & vbCrLf
				Response.Write "g.offset="  & CStr(YOffset) &";" & vbCrLf
				Response.Write "g.title='"  & CStr(Title) &"';" & vbCrLf
				Response.Write "g.xLabel='" & XLabel &"';" & vbCrLf
				Response.Write "g.yLabel='"	& YLabel &"';" & vbCrLf
				Response.Write "g.imageStyle='" & ImageStyle &"';" & vbCrLf
				Response.Write "g.graphStyle='" & Control.Style &"';" & vbCrLf
				If ShowYear Then Response.Write "g.showYear = true;"  & vbCrLf
				If ShowDay Then Response.Write "g.showDay = true;" & vbCrLf
				If XScale>0 Then Response.Write "g.xScale = true; g.s=0;" & vbCrLf 'same effect as using setXScale from javascript. s=0 indicates start X-Axis at 0
				If mLongDate Then Response.Write "g.longDate = true;" & vbCrLf
				If mTScale Then Response.Write "g.setTime(" & mTScaleHour &","& mTScaleMin & ");" & vbCrLf
				If mDScale Then Response.Write "g.setDate(" & mDScaleMon &","& mDScaleDay & ","& mDScaleYear &");" & vbCrLf
				If mLegend Then Response.Write "g.setLegend('" & Join(mLegendData,"','") & "');" & vbCrLf
				'build graph
				Response.Write "g.build();" + vbCrLf
			Response.Write "</script>"
		End Function

		Public Default Function Render()			
			 Dim varStart	 

			 If Control.IsVisible = False Then
				Exit Function
			 End If

			 varStart = Now

			 Call RenderGraph()

			 Page.TraceRender varStart,Now,Control.Name		 	
		End Function
			   	

	End Class
%>

<script language="JavaScript" src="<%=SCRIPT_LIBRARY_PATH%>graph/graph.js"></script>