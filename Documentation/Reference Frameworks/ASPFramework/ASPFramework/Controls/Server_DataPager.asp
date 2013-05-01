<%
	Dim Server_DataPager___
	Server_DataPager___ = 0
	
	Public Function New_ServerDataPager(name) 
		Set New_ServerDataPager = New ServerDataPager
			New_ServerDataPager.Control.Name = name
	End Function
	
	Page.RegisterLibrary "ServerDataPager"
	
	Class ServerDataPager
		Dim Control
		
		Dim VirtualItemCount
		Dim PageIndex
		Dim PageSize
		Dim PagerSize
		Dim PagerType
			
		Dim PrevText
		Dim NextText

		Dim PrevNextCssClass
		Dim PrevNextStyle
		
		Dim PagerStyle
		Dim PagerCssClass
		Dim ShowProgress
		Dim ProgressStyle
		Dim ProgressCssClass
		Dim CurrentPageStyle
		Dim CurrentPageCssClass
		Dim PagerOwner
		
		Private Sub Class_Initialize()
			
			Set Control = New WebControl	
			Set Control.Owner = Me 			
			Set PagerOwner = Nothing

			'AutoName
			Control.Name = "SDP" & Server_DataPager___
			Server_DataPager___ = Server_DataPager___ + 1		

			PagerSize   = 5
			PageSize = 10 'At least
			VirtualItemCount = 0
			PageIndex = 0			
			PagerType = 0
			
			PrevText = "Prev"
			NextText = "Next"						
			PrevNextCssClass = ""
			PrevNextStyle = ""
			CurrentPageCssClass = ""
			CurrentPageStyle = ""
			PagerStyle = ""
			PagerCssClass = ""
			ShowProgress = False
	   End Sub

	   Public Function ReadProperties(bag)
			VirtualItemCount= CInt(bag.Read("A"))
			PageIndex= CInt(bag.Read("B"))
			PageSize= CInt(bag.Read("C"))
			PagerSize= CInt(bag.Read("D"))
			PagerType= CInt(bag.Read("E"))
			PrevText= bag.Read("F")
			NextText= bag.Read("G")
			PrevNextCssClass= bag.Read("H")
			PrevNextStyle= bag.Read("I")
			PagerStyle= bag.Read("J")
			PagerCssClass= bag.Read("K")
			ShowProgress= CBool(bag.Read("L"))
			ProgressStyle= bag.Read("M")
			ProgressCssClass= bag.Read("N")
			CurrentPageStyle= bag.Read("O")
			CurrentPageCssClass= bag.Read("P")
	   End Function
		
	   Public Function WriteProperties(bag)
			bag.Write "A",VirtualItemCount
			bag.Write "B",PageIndex
			bag.Write "C",PageSize
			bag.Write "D",PagerSize
			bag.Write "E",PagerType
			bag.Write "F",PrevText
			bag.Write "G",NextText
			bag.Write "H",PrevNextCssClass
			bag.Write "I",PrevNextStyle
			bag.Write "J",PagerStyle
			bag.Write "K",PagerCssClass
			bag.Write "L",ShowProgress
			bag.Write "M",ProgressStyle
			bag.Write "N",ProgressCssClass
			bag.Write "O",CurrentPageStyle
			bag.Write "P",CurrentPageCssClass
	   End Function
	    
	   Public Function HandleClientEvent(e)
			HandleClientEvent = True
			If Not PagerOwner Is Nothing Then
				PagerOwner.HandleClientEvent(e)
			Else
				Select Case e.EventName
					Case "PageIndexChange"
						Execute "Me.PageIndex = Me.PageIndex + " & e.Instance					
					Case "GotoPageIndex"
						Execute "Me.PageIndex = " & e.Instance					
					Case Else
						HandleClientEvent = ExecuteEventFunctionEX(e)
				End Select				
			End If
	   End Function					
		
	   Private Function RenderNextPrev()
			Dim Pages
			Dim StartIndex,EndIndex
			Dim ProgressMessage			
			If PageSize = 0 Then
				Page.TraceImportantCall Me.Control,"PageSize cannot be zero"
				Exit Function
			End If			
			
			Pages = Int(( (VirtualItemCount-1) / PageSize))
			
			If Pages <= 0 Then
				Response.Write  "<span "
				Response.Write  IIF(ProgressCssClass<>"", " Class='" & ProgressCssClass & "' ","")
				Response.Write  IIF(ProgressStyle<>"", " Style='" & ProgressStyle & "' ","")
				Response.Write ">" & VirtualItemCount & " records found.</span>"
				Exit Function
			End If
			
			StartIndex = PageIndex * PageSize+1
			EndIndex   = StartIndex + PageSize - 1

			If EndIndex > VirtualItemCount Then
				'StartIndex = VirtualItemCount - PageSize + 1
				StartIndex = EndIndex - PageSize+1
				EndIndex = VirtualItemCount				
			End If

			Response.Write  "<span "
			Response.Write  IIF(ProgressCssClass<>"", " Class='" & ProgressCssClass & "' ","")
			Response.Write  IIF(ProgressStyle<>"", " Style='" & ProgressStyle & "' ","")
			Response.Write  ">Records " & StartIndex & " To " & EndIndex & " of " & VirtualItemCount & "</span>&nbsp;"
			
			If PageIndex > 0 Then
				Response.Write "<A "
				Response.Write  IIF(PrevNextCssClass<>"", " Class='" & PrevNextCssClass & "' ","")
				Response.Write  IIF(PrevNextStyle<>"", " Style='" & PrevNextStyle & "' ","")
				Response.Write  Page.GetEventScript("HREF", Control.ControlID, "PageIndexChange", -1, "")
				Response.Write ">"
				Response.Write PrevText
				Response.Write "</A>"				
			End If			
			
			If PageIndex < Pages Then
				If PageIndex>0 Then 
					Response.Write "&nbsp;|&nbsp;"
				End If
				Response.Write "<A "
				Response.Write  IIF(PrevNextCssClass<>"", " Class='" & PrevNextCssClass & "' ","")
				Response.Write  IIF(PrevNextStyle<>"", " Style='" & PrevNextStyle & "' ","")
				Response.Write  Page.GetEventScript("HREF", Control.ControlID, "PageIndexChange",  1, "")
				Response.Write ">"
				Response.Write NextText
				Response.Write "</A>"				
			End If			
			
	   End Function

	   Private Function RenderAsPager()
			Dim Pages,x
			Dim strNav
			Set strNav = New StringBuilder
			Dim strMessage
			Dim StartIndex,EndIndex
			Dim pStart,pEnd,AddAtEnd
		
			If CurrentPageCssClass = "" Then
				CurrentPageCssClass = PagerCssClass
			End If
			
			If CurrentPageStyle = "" Then
				CurrentPageStyle = PagerStyle
			End If

			Pages = Int((VirtualItemCount / PageSize))
			
			If Pages = VirtualItemCount Then
				Pages = VirtualItemCount - 1
			End If
			
			If VirtualItemCount = 0 Then
				strMessage = "No records found."
			ElseIf VirtualItemCount <= PageSize Then
				strMessage = VirtualItemCount & " records found."
			Else
				StartIndex = PageIndex*PageSize+1
				EndIndex   = StartIndex + PageSize - 1
				If EndIndex > VirtualItemCount Then
					StartIndex = EndIndex - PageSize + 1
					EndIndex = VirtualItemCount					
				End If
				strMessage = "<span " &_
					IIF(ProgressCssClass<>"", " Class='" & ProgressCssClass & "' ","") &_
					IIF(ProgressStyle<>"", " Style='" & ProgressStyle & "' ","")  &_
					">" &_
					 StartIndex & " To " & EndIndex & " of " & VirtualItemCount & "</span>"
			End If
			
			strNav.Append strMessage & "&nbsp;&nbsp;&nbsp;"			
			If Pages <=0 Then
				Response.Write strNav.ToString()
				Exit Function
			End If			

			If Pages > PagerSize then
				pStart =  int(PageIndex / PagerSize)* PagerSize
				pEnd   =  pStart + PagerSize-1
				If pStart >0 Then
					strNav.Append "<A " &_
						IIF(PrevNextCssClass<>"", " Class='" & PrevNextCssClass & "' "," ") &_
						IIF(PrevNextStyle<>"", " Style='" & PrevNextStyle & "' "," ")  &_
						Page.GetEventScript("HREF",Control.ControlID,"GotoPageIndex",pStart-1,"") & ">" & PrevText & "</A>"
				End If
				If Pages>pEnd Then
					AddAtEnd =  "<A " &_
					IIF(PrevNextCssClass<>"", " Class='" & PrevNextCssClass & "' "," ") &_
					IIF(PrevNextStyle<>"", " Style='" & PrevNextStyle & "' "," ")  &_
					Page.GetEventScript("HREF",Control.ControlID,"GotoPageIndex",pEnd+1,"") & ">&nbsp;"  & NextText & " </A>"
				End If
				If pEnd>Pages Then 
					pEnd = Pages
				End If
			Else
				pStart = 0
				pEnd   = pages
			End if
			
			For x = pStart To pEnd
				strNav.Append "&nbsp;<A "
				strNav.Append Page.GetEventScript("HREF",Control.ControlID,"GotoPageIndex",x,"") 								
				If PageIndex = x Then
					strNav.Append IIF(CurrentPageCssClass<>"", " Class='" & CurrentPageCssClass & "' ","") 
					strNav.Append IIF(CurrentPageStyle<>"", "Style='" & CurrentPageStyle & "' ","") 
				Else
					strNav.Append IIF(PagerCssClass<>"", " Class='" & PagerCssClass & "' ","") 
					strNav.Append IIF(PagerStyle<>"", " Style='" & PagerStyle & "' ","") 
				End If
				strNav.Append ">" & x & "</A>&nbsp;" 				
			Next
			
			strNav.Append AddAtEnd
			strNav.Append "&nbsp;&nbsp;"
			Response.Write strNav			
	   End Function
		
	   Private Sub RenderAsSuperPager()
		'Use the OnInit...
		'Page X [dropdown] of #pages * Items Per Page [drop down 10/50/100]     [first] next/prev [last]
	   End Sub	
		
	   Public Default Function Render()			
			If PagerType = 0 Then
				RenderNextPrev()
			Else
				RenderAsPager() 				
			End If		
	   End Function
			   	

	End Class
	
%>