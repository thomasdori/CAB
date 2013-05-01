<!--#Include File = "Server_DataPager.asp" -->
<%

	'TEMPLATES	
	Public Sub DataGrid_BlueTemplate(obj,ShowBorder)
		obj.AlternatingItemStyle = "background-color:#DDDDDD;font-size:8pt"
		obj.ItemStyle = "font-size:8pt"
		obj.Control.Style = "border-collapse:collapse"
		obj.HeaderStyle = "font-weight:bold;color:white;background-color:navy;font-size:10pt"
		obj.BorderWidth = IIF(ShowBorder,1,0)
	End Sub

	Public Sub DataGrid_RedTemplate(obj,ShowBorder)
		obj.AlternatingItemStyle = "background-color:#faebd7;font-size:8pt"
		obj.ItemStyle = "font-size:8pt"
		obj.Control.Style = "border-collapse:collapse"
		obj.HeaderStyle = "font-weight:bold;color:white;background-color:#c71585;font-size:10pt"
		obj.BorderWidth = IIF(ShowBorder,1,0)
	End Sub

	Function New_DataGridColumn(ColumnType,DataTextField,DataValueField,DataFormatString)
		Set New_DataGridColumn = New DataGridColumn
		New_DataGridColumn.ColumnType       = ColumnType
		New_DataGridColumn.DataTextField    = DataTextField
		New_DataGridColumn.DataValueField   = DataValueField
		New_DataGridColumn.DataFormatString = DataFormatString
	End Function

	Function New_DataGridTemplateColumn(CellRenderFunctionName)
		Set New_DataGridTemplateColumn = New DataGridColumn
		New_DataGridTemplateColumn.ColumnType = 3
		New_DataGridTemplateColumn.CellRenderFunctionName = CellRenderFunctionName
	End Function

	Page.RegisterLibrary "ServerDataGrid"

	Dim dbgScript
		dbgScript = _ 
		"<script language=""JavaScript"">" + _
		"	function DataGrid_DoScroll(gridID,hdrFixed,scrollPos) {" + _
		"		var elm;" + _
		"		var o,d,t,l;" + _
		"		d = clasp.getObject(gridID+'_scroll');" + _
		"		if (!d) return null;" + _
		"		if (scrollPos) { " + _
		"			var a = scrollPos.split('.');" + _
		"			d.scrollTop = a[0];" + _
		"			d.scrollLeft = a[1];" + _
		"		}" + _
		"		t = parseInt(d.scrollTop);" + _
		"		l = parseInt(d.scrollLeft);" + _
		"		if (hdrFixed) {" + _
		"			o=clasp.getObject(gridID+'_hdr');" + _
		"			if (o) o.style.top=t+'px';" + _
		"			if (o && t==0) o.style.position='relative'; " + _
		"			else o.style.position='absolute'; " + _
		"		}" + _
		"		var elm = clasp.getObject(gridID+'__scrollElm');" + _
		"		if (elm) elm.value = t+'.'+l;" + _
		"	};"
		dbgScript = dbgScript + _
		"	function DataGrid_OnOver(obj,color,cssClass) {" + _
		"		var css = DataGrid_GetCssStyle(obj);" + _
		"		obj.oldColor = css.backgroundColor;" + _
		"		obj.oldCssClass = obj.cssClass;" + _
		"		if(obj && cssClass) obj.className = cssClass;" + _
		"		if(css && color) css.backgroundColor = color;" + _
		"	};" + _
		"	function DataGrid_OnOut(obj,color,cssClass) {" + _
		"		var css = DataGrid_GetCssStyle(obj);" + _
		"		if(!color) color = obj.oldColor;" + _
		"		if(!cssClass) cssClass = obj.oldCssClass;" + _
		"		if(obj && cssClass) obj.className = cssClass;" + _
		"		if(css) css.backgroundColor = color;" + _
		"	};" + _
		"	function DataGrid_GetCssStyle(o){" + _
		"		return (clasp.ua.ns4)? o:o.style;" + _
		"	};" + _
		"</script>"
		Page.RegisterClientStartupScript "DataGridScripts", dbgScript


	Class DataGridRowContext
		Private mCssName
		Private mStyle
		Private mCaption
		Private mCaptionStyle
		Private mCaptionCssName
		Private mCursor
		Private mAttributes
		Public DataSource
		Public WasAltered
		Public Mode
		
		Private Sub Class_Initialize()
			InitContext Nothing
		End Sub
		
		Public Sub InitContext(ds)
			Set DataSource = ds
			mStyle = ""
			mCssName = ""
			mAttributes = ""
			mCursor = False
			WasAltered = false
		End Sub
		
		Public Sub Reset() 
			mStyle = ""
			mCssName = ""
			mAttributes = ""
			mCaption = ""			
			mCursor = False
			WasAltered = false
		End Sub

		Public Property Get Caption
			Caption = mCaption
		End Property

		Public Property Let Caption(value)
			mCaption= value
			WasAltered = True
		End Property

		Public Property Get CaptionStyle
			CaptionStyle = mCaptionStyle
		End Property

		Public Property Let CaptionStyle(value)
			mCaptionStyle= value
			WasAltered = True
		End Property

		Public Property Get CaptionCssName
			CaptionCssName = mCaptionCssName
		End Property

		Public Property Let CaptionCssName(value)
			mCaptionCssName= value
			WasAltered = True
		End Property


		Public Property Get CursorHighLighting
			CursorHighLighting = mCursor
		End Property

		Public Property Let CursorHighLighting(value)
			mCursor= value
			WasAltered = True
		End Property
		
		Public Property Get CssName
			CssName = mCssName
		End Property

		Public Property Let CssName(value)
			mCssName= value
			WasAltered = True
		End Property

		Public Property Get Style
			Style = mStyle			
		End Property

		Public Property Let Style(value)
			mStyle = Value
			WasAltered = True
		End Property
		
		Public Property Get Attributes
			Attributes= mAttributes			
		End Property

		Public Property Let Attributes(value)
			mAttributes= Value
			WasAltered = True
		End Property

	End Class
	
	
	'Special Variables used with Header Column Spanning
	Dim mColSpan:mColSpan = 0
	Dim mColSpanCount:mColSpanCount = 0
	
	Class DataGridColumn
	
		Dim ColumnType   
				'1 Bound, 2 Format, 3 Template, 4 CheckBox, 5 Radio, 6 'Edit Column
			    'Extended - Private : 100 (If DataFormatFunctionName is supplied)
		Dim CellRenderFunctionName
		Dim HorizonalAlign
		Dim VerticalAlign
		Dim ColumnWidth
		Dim DataTextField
		Dim DataValueField
		Dim HeaderText
		Dim Visible
		Dim SortColumn
		Dim SortOrder	'0 - None, 1- ASC, 2 - DESC
		Dim OrdinalPosition
		Dim NoWrap
		Dim CssClass
		Dim Style
		Dim RowSpan
		Dim HeaderSpan
		dim Locked		

		Dim ReadOnly
		Dim EditControl
		Dim DataFormatString
		Dim DataFormatFunctionName
		Dim DataSource
		Dim DataGridOwner
		
		Private fDataTextField
		Private fDataValueField
		Private fncCellRenderFunction
		Private fncDataFormatFunctionName
	
		Private mSelectedValue
		Private mRenderStart
		Private mRenderEnd
		Private mAltRenderStart		
		Private mRowSpanCount
		Private mHdrSpanCount
		Private Sub Class_Initialize()
			
			OrdinalPosition = 0
			ColumnType		= 1
			Visible			= True
			ReadOnly		= True		
			NoWrap			= False
			CssClass		= ""
			Style			= ""
			SortOrder		= 0
 			RowSpan			= 0
 			HeaderSpan		= 0
 			mRowSpanCount	= 0
 			mHdrSpanCount	= 0
 			Locked	        = False
			Set EditControl = Nothing
			Set DataSource  = Nothing
			Set fncCellRenderFunction = Nothing
			Set fncDataFormatFunctionName = Nothing
			Set fDataTextField  = Nothing
			Set fDataValueField = Nothing
			Set DataGridOwner   = Nothing			
		End Sub
	
		Private Function GetTDTag(Style,CssClass,IsHeader)
 			Dim mHdrSpan
 			Dim mRowSpan
 			mHdrSpan = IIf(IsHeader And HeaderSpan>0, HeaderSpan, 0)
 			mRowSpan = IIf(Not IsHeader And RowSpan>0, RowSpan, 0)
 			
 			If Locked Then
 				CssClass = "locked"
 			End If
 			
			GetTDTag = "<td " 			
				If HorizonalAlign <> ""	Then GetTDTag = GetTDTag & " align='" & HorizonalAlign & "' "
				If VerticalAlign  <> "" Then GetTDTag = GetTDTag & " valign='" & VerticalAlign & "' "
				If Style		  <> ""	Then GetTDTag = GetTDTag & " style='" & Style & "' "
				If ColumnWidth    <> ""	Then GetTDTag = GetTDTag & " width='" & ColumnWidth & "' "
				If CssClass	      <> "" Then GetTDTag = GetTDTag & " class='" & CssClass & "' "
 				If mRowSpan		  <> 0  Then GetTDTag = GetTDTag & " rowspan='" & mRowSpan & "' "
 				If mHdrSpan		  <> 0  Then GetTDTag = GetTDTag & " colspan='" & mHdrSpan & "' "
				If NoWrap			    Then GetTDTag = GetTDTag & " nowrap "
			GetTDTag = GetTDTag & ">"
		End Function
		
		Public Property Get IsChecked()
			If ColumnType = 5 Then
				IsChecked = (Request.Form(DataGridOwner.Control.ControlID & "_" & OrdinalPosition) <> "")
			End If
		End Property
		
		Public Property Get GetCheckedValue()
			If ColumnType = 5 Then
				GetCheckedValue = Request.Form(DataGridOwner.Control.ControlID & "_" & OrdinalPosition)
			End If			
		End Property
		
		Private Sub Init()
			mRenderStart = GetTDTag(Style,CssClass,False)				
			
			If DataTextField<>"" Then
				Set fDataTextField = DataSource.Fields(DataTextField)
			End If
			
			If  DataValueField <> "" Then
				Set fDataValueField = DataSource.Fields(DataValueField)
			End If
				
			If CellRenderFunctionName <> "" Then
				Set fncCellRenderFunction = GetFunctionReference(CellRenderFunctionName)
			End If				
		
			If DataFormatFunctionName <> "" Then
				Set fncDataFormatFunctionName = GetFunctionReference(DataFormatFunctionName)
				ColumnType = 100 'To make it faster
			End If			
						
			mRenderEnd = "</td>"
			
			ReadOnly = (EditControl Is Nothing)
			
		End Sub
		
		Public Sub Render(Mode) 'mode = 1 normal, 2 = selected, 3 = edit
			Dim keyField
			
			'row span check
 			If RowSpan>0 Then 
 				If (mRowSpanCount mod RowSpan) <> 0 Then 
	 				mRowSpanCount = mRowSpanCount+1
 					Exit Sub
 				ElseIf mRowSpanCount = 0 Then
 					mRowSpanCount = 1
 				End If
 			End If
 			If mRowSpanCount > 1 Then mRowSpanCount = 1
			
			If mRenderStart = "" Then
				Call Init()
			End If
			
			Response.Write mRenderStart
						
			If Mode = 3 And Not ReadOnly Then
				Dim e
				Set e = Page.GetPostBackEvent()
				If e.TargetObject Is DataGridOwner.Control And e.EventName = "Edit" Then				
					EditControl.SetValueFromDataSource(fDataTextField.Value)
				End If
				Call EditControl.Render()
			Else	
				Select Case ColumnType
					Case 1:
						'Response.Write Server.HTMLEncode( "" & fDataTextField.Value )
						Response.Write fDataTextField.Value
					Case 2:
						Response.Write Replace(Replace( Replace(DataFormatString,"{R}",DataSource.AbsolutePosition) ,"{0}",fDataTextField.Value), "{1}",fDataValueField)
					Case 3:
						If Not fncCellRenderFunction Is Nothing Then
							Call fncCellRenderFunction(DataSource)
						End If
					Case 100:
						If Not fncDataFormatFunctionName Is Nothing Then
							Response.Write fncDataFormatFunctionName(fDataTextField.Value)
						End If
					Case 4:
						Response.Write "<input type='checkbox' value='" & Server.HTMLEncode(fDataValueField.Value) & "' name='" & DataGridOwner.Control.ControlID & "_" & DataSource.AbsolutePosition  & "_" & OrdinalPosition & "'> " & fDataTextField.Value
					Case 5:
						Response.Write "<input type='radio' " &  IIF(fDataValueField.value = mSelectedValue,"CHECKED","")  & " value='" & Server.HTMLEncode(fDataValueField.Value) & "' name='" & DataGridOwner.Control.ControlID & "_" & OrdinalPosition & "'> " & fDataTextField.Value
					Case 6:
						If Mode = 3 Then
							If DataGridOwner.DataKeyFieldName = "" Then
								keyField = DataSource.AbsolutePosition								
							Else
								keyField = DataGridOwner.DataKeyField
							End If
							Response.Write " <A " 'style='color:white' " 
								Response.Write Page.GetEventScript("HREF",DataGridOwner.Control.ControlID,"Cancel",KeyField,"")
								Response.Write " >Cancel</a> | <A " 'style='color:white' " 
								Response.Write Page.GetEventScript("HREF",DataGridOwner.Control.ControlID,"Update",KeyField,"")
							Response.Write " >Update</a>"		 			
						Else
							Response.Write " <A " 'style='color:black' " 
							Response.Write Page.GetEventScript("HREF",DataGridOwner.Control.ControlID,"Edit",DataSource.AbsolutePosition,"")
							Response.Write " >Edit</a>"
						End If
				End Select
			End If						
			Response.Write mRenderEnd
		End Sub
		
		Public Sub RenderHeader()
			Dim t
			
			'header span check
			If HeaderSpan>0 And mColSpan=0 Then mColSpan = HeaderSpan
 			If mColSpan>0 Then 
 				If (mColSpanCount mod mColSpan) <> 0 Then 
	 				mColSpanCount = mColSpanCount+1
 					Exit Sub
 				ElseIf mColSpanCount =0 Then
 					mColSpanCount	= 1
 				End If
 			End If
 			If mColSpanCount > 1 Then 
 				mColSpan		= 0
 				mColSpanCount	= 0
 			End If
 			
			Response.Write GetTDTag(DataGridOwner.HeaderStyle,DataGridOwner.HeaderCssClass,True)
			If SortColumn<>"" Then
				t = "Sort this Column"
				'check for html inside header text
				If Not(InStr(1,HeaderText,"<")>0 And InStr(1,HeaderText,">")>0) Then
					t = "Sort by "+Replace(HeaderText,"'","")
				End If
				Response.Write "<a title='"+ t +"' "
					Response.Write IIF(DataGridOwner.SortStyle<>"", " style = '" & DataGridOwner.SortStyle & "' ","")
					Response.Write IIF(DataGridOwner.SortCssClass<>"", " class = '" & DataGridOwner.SortCssClass & "' ","")
				    Response.Write Page.GetEventScript("href",DataGridOwner.Control.ControlID,"ColumnSort",SortColumn,"") & ">"
				    Response.Write IIF(SortOrder,"<image src='"+SCRIPT_LIBRARY_PATH+"datagrid/images/"+IIf(SortOrder=1,"up","dn")+".gif' width='8' height='9' border='0' align='right'>","")
				    Response.Write HeaderText & "</a>"
			Else
				Response.Write HeaderText
			End If
			Response.Write "</td>"
		End Sub	
		
	End Class

	Public Function New_ServerDataGrid(name) 
		Set New_ServerDataGrid = New ServerDataGrid
			New_ServerDataGrid.Control.Name = name
	End Function

	 Class ServerDataGrid
		
		Dim Control
		Dim Columns
		Dim DataSource
		
		Dim ShowHeader					
		Dim HorizonalAlign
		Dim CellPadding
		Dim CellSpacing
		Dim BorderWidth
		Dim BorderColor
		Dim BackImageURL			
		Dim TableStyle
		Dim TableCssClass
		Dim TableWidth	
		Dim TableHeight
		Dim AllowCustomPaging
		Dim AllowPaging
		Dim PagerStyle
		Dim PagerCssClass
		Dim AutoGenerateColumns
		Dim CursorHighlighting
		Dim EnableScrollBars
		Dim FixedColumnHeader
					
		Dim HeaderStyle
		Dim FooterStyle
		
		Dim ItemStyle
		Dim AlternatingItemStyle
		Dim EditItemStyle
		Dim SelectedItemStyle
		Dim SortStyle
		
		Dim HeaderCssClass			
		Dim FooterCssClass
		Dim ItemCssClass
		Dim AlternatingItemCssClass
		Dim EditItemCssClass
		Dim SelectedItemCssClass	
		Dim SortCssClass	

		Dim ItemHighlightColor
		Dim ItemHighlightCssClass
		Dim AlternatingHighlightColor
		Dim AlternatingHighlightCssClass
		
		Dim EditItemIndex
		Dim SelectedItemIndex
		Dim SelectedItemKeyValue

		Dim HeaderFunctionName
		Dim FooterFunctionName
		
		Dim Pager
		Dim PagerHorizontalAlign
		Dim PagerShowOnTop
		
		Dim DataKeyField
		Dim DataKeyFieldName
		
		Dim CurrentRow,CurrentColumn
		
		'To Alter the Row
		Dim RowDataBound
		
		Private mScrollPos 'stores  scroll position
		
		Private Sub Class_Terminate()
			Set DataSource = Nothing
		End Sub

		
		Private Sub Class_Initialize()

			Set Control =  New WebControl
			Set Control.Owner = Me			
			Control.ImplementsProcessPostBack = True

			Set DataSource = Nothing

			ShowHeader = True
			HorizonalAlign = ""
			CellPadding = 2
			CellSpacing = 0
			BorderWidth = 0
			BorderColor = ""
			BackImageURL = ""
			
			EnableScrollBars			 = False
			FixedColumnHeader			 = False
			CursorHighlighting			 = False
			ItemHighlightColor			 = "#B5BDD6"
			ItemHighlightCssClass		 = ""
			AlternatingHighlightColor	 = "#B5BDD6"
			AlternatingHighlightCssClass = ""
			
			TableStyle = ""
			TableCssClass = ""

			AllowCustomPaging = False
			AllowPaging       = False
			AutoGenerateColumns = True

			HeaderStyle = ""
			EditItemStyle = ""
			SelectedItemStyle = ""
			HeaderCssClass = ""
			EditItemCssClass = ""
			SelectedItemCssClass = ""	

			EditItemIndex = -1
			SelectedItemIndex = -1
			SelectedItemKeyValue = ""

			HeaderFunctionName = ""
			FooterFunctionName = ""			 

			CurrentRow = 0
			CurrentColumn  = 0

			RowDataBound = ""
			DataKeyFieldName = ""
			Set DataKeyField = Nothing
			mScrollPos = "0.0"

			Set Pager = New ServerDataPager
			PagerHorizontalAlign = "right"
			PagerShowOnTop = True
			 
		End Sub

		Public Sub WriteProperties(bag)
			bag.Write "A",ShowHeader
			bag.Write "B",HorizonalAlign
			bag.Write "C",CellPadding
			bag.Write "D",CellSpacing
			bag.Write "E",BorderWidth
			bag.Write "F",BorderColor
			bag.Write "G",BackImageURL
			bag.Write "H",TableStyle
			bag.Write "I",TableCssClass
			bag.Write "J",AllowCustomPaging
			bag.Write "K",AllowPaging
			bag.Write "L",AutoGenerateColumns
			bag.Write "M",HeaderStyle
			bag.Write "N",FooterStyle
			bag.Write "O",ItemStyle
			bag.Write "P",AlternatingItemStyle
			bag.Write "Q",EditItemStyle
			bag.Write "R",SelectedItemStyle
			bag.Write "S",HeaderCssClass
			bag.Write "T",FooterCssClass
			bag.Write "U",ItemCssClass
			bag.Write "V",AlternatingItemCssClass
			bag.Write "W",EditItemCssClass
			bag.Write "X",SelectedItemCssClass
			bag.Write "Y",EditItemIndex
			bag.Write "Z",SelectedItemIndex
			bag.Write "Z0",SelectedItemKeyValue
			bag.Write "Z1",HeaderFunctionName
			bag.Write "Z2",FooterFunctionName
			bag.Write "Z3",PagerHorizontalAlign
			bag.Write "Z4",PagerShowOnTop
			bag.Write "Z5",TableWidth
			bag.Write "Z6",SortStyle
			bag.Write "Z7",SortCssClass
			bag.Write "Z8",DataKeyFieldName
			bag.Write "Z9",TableHeight
			bag.Write "C1",CursorHighlighting
			bag.Write "C2",ItemHighlightColor
			bag.Write "C3",ItemHighlightCssClass
			bag.Write "C4",AlternatingHighlightColor
			bag.Write "C5",AlternatingHighlightCssClass
			bag.Write "D1",EnableScrollBars
			bag.Write "D2",FixedColumnHeader
			bag.Write "D3",mScrollPos
		
			'Persist Column Settings
			'For Each col in Columns
			' bag.Write "CNT",UBOUND(Columns)
			' bag.Write "C" & X & "SH",col.ShowHeader
			'	
			'Next

		End Sub
	
		Public Sub ReadProperties(bag)
			ShowHeader= CBool(bag.Read("A"))
			HorizonalAlign= bag.Read("B")
			CellPadding= bag.Read("C")
			CellSpacing= bag.Read("D")
			BorderWidth= bag.Read("E")
			BorderColor= bag.Read("F")
			BackImageURL= bag.Read("G")
			TableStyle= bag.Read("H")
			TableCssClass= bag.Read("I")
			AllowCustomPaging= CBool(bag.Read("J"))
			AllowPaging= CBool(bag.Read("K"))
			AutoGenerateColumns= CBool(bag.Read("L"))
			HeaderStyle= bag.Read("M")
			FooterStyle= bag.Read("N")
			ItemStyle= bag.Read("O")
			AlternatingItemStyle= bag.Read("P")
			EditItemStyle= bag.Read("Q")
			SelectedItemStyle= bag.Read("R")
			HeaderCssClass= bag.Read("S")
			FooterCssClass= bag.Read("T")
			ItemCssClass= bag.Read("U")
			AlternatingItemCssClass= bag.Read("V")
			EditItemCssClass= bag.Read("W")
			SelectedItemCssClass= bag.Read("X")
			EditItemIndex= CInt(bag.Read("Y"))
			SelectedItemIndex= CInt(bag.Read("Z"))
			SelectedItemKeyValue = bag.Read("Z0")
			HeaderFunctionName= bag.Read("Z1")
			FooterFunctionName= bag.Read("Z2")
			PagerHorizontalAlign= bag.Read("Z3")
			PagerShowOnTop= CBool(bag.Read("Z4"))
			TableWidth = bag.Read("Z5")
			SortStyle = bag.Read("Z6")
			SortCssClass = bag.Read("Z7")
			DataKeyFieldName = bag.Read("Z8")
			TableHeight = bag.Read("Z9")
			CursorHighlighting = bag.Read("C1")
			ItemHighlightColor = bag.Read("C2")
			ItemHighlightCssClass = bag.Read("C3")
			AlternatingHighlightColor = bag.Read("C4")
			AlternatingHighlightCssClass = bag.Read("C5")
			EnableScrollBars = bag.Read("D1")
			FixedColumnHeader = bag.Read("D2")
			mScrollPos = bag.Read("D3")
		End Sub
		
		Public Function ProcessPostBack()
			Dim key:key = Control.ControlID &"__scrollElm"	
			If Request.Form(key) <> "" Then
				mScrollPos = Request.Form(key)
			End If		
		End Function		
	
		Public Sub OnInit()			
			  Pager.Control.ControlID = Control.ControlID & "_Pager"
			  'Set Pager.Control.Parent = Me  'If you want to make it 100% dependant... no need though...
			  Set Pager.PagerOwner = Me
		End Sub

		Public Function HandleClientEvent(e)
			e.Source = Me.Control.ControlID
			Select Case e.EventName
				Case "PageIndexChange"					
					HandleClientEvent = ExecuteEventFunctionEX(e)
					Pager.PageIndex = Pager.PageIndex + CInt(e.Instance)
				Case "GotoPageIndex"					
					e.EventName = "PageIndexChange"
					HandleClientEvent = ExecuteEventFunctionEX(e)
					Pager.PageIndex = CInt(e.Instance)
				Case Else
					HandleClientEvent = ExecuteEventFunctionEX(e)
			End Select				
	    End Function					
	
		'Renders the pager and return the max. number of rows to browse for
		Private Function RenderPager()								
			RenderPager = Pager.PageSize							
			If 	AllowCustomPaging Then
				If Pager.VirtualItemCount = 0 Then
					Pager.VirtualItemCount = DataSource.RecordCount
				End If
			Else
				Pager.VirtualItemCount = DataSource.RecordCount
				If DataSource.RecordCount > Pager.PageSize  Then '0  Then 
					DataSource.AbsolutePosition = (Pager.PageIndex * Pager.PageSize) + 1				
				End If
			End If			
			Pager.Render						
		End Function

		Public Function CreateColumns(arrColumnNames,arrHeaders)
			Dim x,mx
			Dim bolHeaderPresent
			
			bolHeaderPresent =  IsArray(arrHeaders) 
						
			If IsArray(Columns) Then
				Erase Columns
			End If
			
			Redim Columns(UBound(arrColumnNames))
			mx = UBound(Columns)
			For x = 0 To mx
				Set Columns(x) = New DataGridColumn
				Columns(x).OrdinalPosition = x
				Columns(x).DataTextField = arrColumnNames(x)
				If bolHeaderPresent Then
					If arrHeaders(x) <> "" Then
						Columns(x).HeaderText = arrHeaders(x)
					Else
						Columns(x).HeaderText = arrColumnNames(x)
					End If
				Else
					Columns(x).HeaderText = arrColumnNames(x)
				End If	
			Next
		End Function

		
		Public Function SetGridColumns(value)
			Dim x,col
			Columns  = value
			x = 0
			For Each col in Columns
				col.OrdinalPosition = x
				x= x +  1
			Next
		End Function
		
		Public Property Get ColumnsCount()
			If IsArray(Columns) Then
				ColumnsCount  = UBound(Columns)
			Else
				ColumnsCount  = 0
			End If
		End Property
			
		Public Sub GenerateColumns()
			Dim fld
			Dim x
			
			If IsArray(Columns) Then
				Erase Columns
			End If
			
			Redim Columns(DataSource.Fields.Count-1)

			x=0
			For Each fld in DataSource.Fields
				Set Columns(x) = New DataGridColumn
				With Columns(x)
					 Set .DataGridOwner = Me
					 Set .DataSource	= DataSource
					 .OrdinalPosition	= x
					.DataTextField		= fld.Name
					.HeaderText			= fld.Name										
				End With			
				x=x+1
			Next
		End Sub
				
		Public Function GetColumnByName(Name)
			Dim col
			
			Set GetColumnByName  = Nothing
						
			For Each col in Columns
				If col.DataTextField = Name Then
					Set GetColumnByName = col
					Exit For
				End If
			Next
		
		End Function
				
		Public Default Function Render()			
			Dim col
			Dim scrlTop
			Dim varStart
			Dim alt
			Dim mode
			Dim maxRows
			Dim RowTag,AltRowTag,EditRowTag,SelectedRowTag
			Dim TheRowTag
			Dim fncRowDataBound
			Dim objRowContext
			Dim blnUseRowContext
			Dim fnc
			Dim InstanceValue
			Dim CursorAttributes
			Dim ItmCursorAttributes
			Dim AltCursorAttributes
									
			varStart = Now
			
			If Control.Visible = False Then
				Exit Function
			End If
				
			If DataSource Is Nothing Then
				Pager.VirtualItemCount = 0
				Pager.PageIndex = 0
				'Pager.Pages = 0
				Exit Function
			End If
						
			If RowDataBound <> "" Then
				Set fncRowDataBound = GetFunctionReference(RowDataBound)
				Set objRowContext = New DataGridRowContext
				objRowContext.InitContext DataSource
				blnUseRowContext = True
			Else
				Set fncRowDataBound = Nothing
				blnUseRowContext = False
			End If
						
			If AutoGenerateColumns Then				
				GenerateColumns()
			Else
				CurrentColumn = 0
				For Each col in Columns
					Set col.DataSource = DataSource
					Set col.DataGridOwner = Me
					CurrentColumn = CurrentColumn  +  1
				Next
			End If
			
			If Columns(0).Locked Then
				Response.Write "<LINK rel='stylesheet' type='text/css' href='" & SCRIPT_LIBRARY_PATH & "datagrid/GridColFreeze.css'>"
			End If
			
			If DataKeyFieldName <> "" Then
				Set DataKeyField = DataSource(DataKeyFieldName)
			End If
									
			alt = 1

			If Columns(0).Locked Then
				Response.Write "<div id='div-datagrid'>"
			End If			
			
			If EnableScrollBars Then
				Response.write "<div id='"+Control.ControlID+"_scroll' style='overflow:auto;clip:auto;"
				Response.write IIF(TableWidth<>"","width:" & TableWidth &";","")
				Response.write IIF(TableHeight<>"","height:" & TableHeight &";","")
				Response.write "border:1 solid #c0c0c0' "
				Response.write "onscroll=""DataGrid_DoScroll('"& Control.ControlID &"',"& IIf(FixedColumnHeader,"true","false") &")"""
				Response.write "><input type='hidden' id='"+Control.ControlID+"__scrollElm' name='"+Control.ControlID+"__scrollElm' value='"& mScrollPos &"'>"
			End IF
			
			Response.Write  vbNewLine &  "<table cellspacing=" & CellSpacing & " cellpadding=" & CellPadding
			Response.Write " id='" &  Control.ControlID & "' "
			Response.Write IIF(Control.CssClass<>""," class='" & Control.CssClass & "' ","")
			Response.Write IIF(Control.Style<>""," style='" & Control.Style & "' ","")
			Response.Write IIF(HorizonalAlign<>""," align=" & HorizonalAlign ,"")
			Response.Write IIF(TableWidth<>""," width =" & TableWidth,"")
									
			If BorderWidth > 0 Then
				Response.Write " border=" & BorderWidth & " bordercolor='" & BorderColor & "'"
			End If

			Response.Write ">" & vbNewLine

			If HeaderFunctionName<> "" Then
				Response.Write "<tr><td colspan=" & Ubound(Columns) + 1 & " >"
				Set fnc = GetFunctionReference(HeaderFunctionName)
				If Not fnc Is Nothing Then
					Call fnc()
				End If
				Response.Write "</td></tr>"
			End If

			If AllowPaging Then				
				Response.Write "<tr><td colspan=" & Ubound(Columns) + 1 & " " & IIF(PagerStyle<>"","style='" & PagerStyle & "' ","") & " " & IIF(PagerCssClass<>"","class='" & PagerCssClass & "' ","") 
				Response.Write " align='" & PagerHorizontalAlign & "' >"
				maxRows   =  RenderPager() 				
				Response.Write "</td></tr>"
			Else
				If DataSource.ActiveConnection Is Nothing Then
					maxRows = DataSource.RecordCount
				Else
					maxRows = 100000000
				End If
			End If						
			
			If ShowHeader Then				

				If IsNumeric(mScrollPos) Then scrlTop = CInt(mScrollPos)
				Response.Write "<tr id='"+Control.ControlID+"_hdr' "+IIf(scrlTop>0 And FixedColumnHeader,"style='position:absolute;top:"& scrlTop &"'","")+">"
				CurrentColumn = 0
				For Each col In Columns					
					If col.Visible Then
						col.RenderHeader()					
					End If
				Next
				Response.Write "</tr>"
			End If			

			If CursorHighlighting Then 
				CursorAttributes = Page.GetEventScript("onclick",Control.ControlID,"OnClick","{0}","")
				ItmCursorAttributes = CursorAttributes + " onmouseover=""DataGrid_OnOver(this,'"+ItemHighlightColor+"','"+ItemHighlightCssClass+"');"" onmouseout=""DataGrid_OnOut(this);"" "
				AltCursorAttributes = CursorAttributes + " onmouseover=""DataGrid_OnOver(this,'"+AlternatingHighlightColor+"','"+AlternatingHighlightCssClass+"');"" onmouseout=""DataGrid_OnOut(this);"" "
			End If

			RowTag		   = "<tr " & IIF(ItemStyle<>"","style='" & ItemStyle & "' ","") & IIF(ItemCssClass<>"","class='" & ItemCssClass & "' ","") & ItmCursorAttributes & ">"
			AltRowTag	   = "<tr " & IIF(AlternatingItemStyle<>"","style='" & AlternatingItemStyle & "' ","") & IIF(AlternatingItemCssClass<>"","class='" & AlternatingItemCssClass & "' ","") & AltCursorAttributes & ">"
			EditRowTag	   = "<tr " & IIF(EditItemStyle<>"","style='" & EditItemStyle & "' ","") & IIF(EditItemCssClass<>"","class='" & EditItemCssClass & "'","") & ">"
			SelectedRowTag = "<tr " & IIF(SelectedItemStyle<>"","style='" & SelectedItemStyle & "' ","") & IIF(SelectedItemCssClass<>"","class='" & SelectedItemCssClass & "'","") & ">"
						
			CurrentRow = 0		
			While Not DataSource.EOF And CurrentRow < maxRows
			
				Response.Write vbNewLine

				mode = alt							
				If DataSource.AbsolutePosition = EditItemIndex Then
					mode = 3
				ElseIf DataSource.AbsolutePosition = SelectedItemIndex Then
					mode = 2
				End If
				If DataKeyFieldName<>"" Then
					If DataSource(DataKeyFieldName)&"" =SelectedItemKeyValue &"" Then
						mode = 2
					End If
				End If
				
				Select Case Mode
					Case 0: TheRowTag = RowTag
					Case 1: TheRowTag = AltRowTag
					Case 2: TheRowTag = SelectedRowTag
					Case 3: TheRowTag = EditRowTag
					Case Else: TheRowTag = RowTag
				End Select
				
				If blnUseRowContext Then
					'setup default row properties
					objRowContext.Mode = Mode
					objRowContext.Style = Array(ItemStyle,AlternatingItemStyle,EditItemStyle,SelectedItemStyle)(Mode)
					objRowContext.CssName = Array(ItemCssClass,AlternatingItemCssClass,EditItemCssClass,SelectedItemCssClass)(Mode)
					objRowContext.CursorHighlighting = CursorHighlighting
					objRowContext.WasAltered = False
					'call RowDataBound function
					fncRowDataBound objRowContext
					If objRowContext.WasAltered Then
						CursorAttributes = ""
						If objRowContext.CursorHighlighting Then
							If Mode = 0 Then CursorAttributes = ItmCursorAttributes
							If Mode = 1 Then CursorAttributes = AltCursorAttributes
						End If
						TheRowTag =  "<tr " & IIF(objRowContext.Style<>"","style='" & objRowContext.Style & "' ","") & IIF(objRowContext.CssName <>"","class='" & objRowContext.CssName & "' ","") & IIF(objRowContext.Attributes<>"",objRowContext.Attributes,CursorAttributes) & ">"
						If objRowContext.Caption <> "" Then
							TheRowTag = "<tr><td "+ _
								IIf(objRowContext.CaptionCssName<>""," class='" & objRowContext.CaptionCssName &"' ","") + _
								IIf(objRowContext.CaptionStyle<>""," style='" & objRowContext.CaptionStyle &"' ","") + _
								" colspan='" & Ubound(Columns) + 1 & "'>" & objRowContext.Caption &"</td></tr>" + TheRowTag
						End If
						objRowContext.Reset
					End If
				End If

				'Write TheRowTag
				If CursorHighlighting Then 
					If DataKeyFieldName<>"" Then 
						InstanceValue = Replace(DataSource(DataKeyFieldName)&"","'","\'")
					Else
						InstanceValue = DataSource.AbsolutePosition
					End If
					Response.Write Replace(TheRowTag,"{0}",InstanceValue)
				Else
					Response.Write TheRowTag
				End If

				CurrentColumn = 0

				For Each col In Columns
					If col.Visible Then
						col.Render(mode)
					End If
					CurrentColumn = CurrentColumn + 1
				Next
				Response.Write "</tr>"
				alt = 1 - alt
				DataSource.MoveNext
				CurrentRow = CurrentRow + 1
			Wend

			If FooterFunctionName<> "" Then
				Response.Write "<tr><td colspan=" & Ubound(Columns) + 1 & ">"
				Set fnc = GetFunctionReference(FooterFunctionName)
				If Not fnc Is Nothing Then
					Call fnc()
				End If
				Response.Write "</td></tr>"
			End If
							
			Response.Write "</table>" & vbNewLine
			If EnableScrollBars Then
				Response.Write "</div>"
				Response.Write "<script>DataGrid_DoScroll('"& Control.ControlID &"',null,'"& mScrollPos &"')</script>"
			End If
			
			If Columns(0).Locked Then
				Response.Write "</div>"
			End If			

			Page.TraceRender varStart,Now,Control.ControlID
			
		End Function
	
	End Class
%>