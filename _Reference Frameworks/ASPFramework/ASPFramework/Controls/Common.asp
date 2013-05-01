<%

Public Function GetArrayDimensions(arr)
	Dim x
	On Error Resume Next
	
	If Not IsArray(arr) Then
		GetArrayDimensions = 0
		Exit Function
	End If
	
	For x = 1 To 10 'Max 10 Dimensions
		GetArrayDimensions = UBound(arr,x)
		If Err.number > 0 Then
			Err.Clear				
			Exit For
		End If				
	Next
	GetArrayDimensions = x - 1
End Function

Public Function ConvertToRecordSet(DataSource)
	Dim objRs

	Set objRs = Server.CreateObject("ADODB.Recordset")
	objRs.CursorLocation = 3 'adUseClient
	objRs.CursorType     = 1 'adOpenForwardOnly
	objRs.LockType       = 4 'adLockBatchOptimistic
	
	Select Case TypeName(DataSource)
		Case "Dictionary"
				Dim sKey
				objRs.Fields.Append "Col0", 200, 50
				objRs.Fields.Append "Col1", 200, 50				 
				objRs.Open
				For Each sKey In DataSource.Keys
					objRs.AddNew
					objRs(0).Value = sKey
					objRs(1).Value = CStr(DataSource.Item(sKey))
				Next
				
			Case "Variant()" ' An array was passed in
					Dim Dims
					Dim c1,c2
					Dim r1,r2					
					Dim col,row
					
					If IsArray(DataSource) Then
						Dims = GetArrayDimensions(DataSource)
						If Dims>2 Then
							'NO NO NO
							Err.Raise vbError + 10000,"ConvertToRecordSet","Array can only have up to 2 dimensions"
						End If

						r1 = LBound(DataSource, 1)
						r2 = UBound(DataSource, 1)						
						'Accept only Array(ROW,COLS)						
						If Dims > 1 Then
							c1 = LBound(DataSource, 2)
							c2 = UBound(DataSource, 2)
							For col = c1 To c2
								objRs.Fields.Append "Col" & col , 200, 50
							Next
							objRs.Open
							For row = r1 To r2
								objRs.AddNew
								For col = c1 To c2
									objRs("Col" & col).Value = CStr(DataSource(row, col))
								Next
							Next
						Else
							objRs.Fields.Append "Col0", 200, 50 
							objRs.Open
							For row = r1 To r2
								objRs.AddNew
								objRs(0).Value = CStr(DataSource(row))
							Next
						End If
					End If
				Case Else
					Err.Raise vbError + 10000,"ConvertToRecordSet","Object Type " & TypeName(DataSource) & " Not Supported!"
					Exit Function
		End Select		
		objRs.Update
		objRs.MoveFirst		
		Set ConvertToRecordSet = objRs		
End Function

'FOR USING WITH CHECKBOXES, RADIO BUTTONS, DROPDOWNS/LISTBOXES!
Public Function CacheListItemsCollection(objRs,DataTextField,DataValueField,CacheName,bolAddBlank,DefaultValue)
						
			Dim x,mx	
			Dim fld1,fld2
			Dim use1,use2
			Dim Items

			Set Items = New_ListItemsCollectionObject()
						
			If DataTextField<>""  Then 
				Set fld1 = objRs(DataTextField) 
				use1 = True
			Else  
				Set fld1 = Nothing
				use1 = False
			End If
			
			If DataValueField<>"" Then 
				Set fld2 = objRs(DataValueField) 
				use2 = True
			Else 
				Set fld2 = Nothing
				use2 = false
			End If
			
			x=0
			While Not objRs.EOF									
				Items.Append IIF(use1, fld1,""),  IIF(use2, fld2,x),False
				objRs.MoveNext
				x = x + 1		
			Wend
			
			If bolAddBlank Then
				Items.Add "","",True,-1
			End If			
			
			If DefaultValue <> "" Then
				Items.SetSelectedByValue DefaultValue,True
			End If
			
			Set fld1 = Nothing
			Set fld2 = Nothing
						
			Application.Lock
				Application(CacheName) = Items.GetState()
				Application.UnLock
			Err.Clear

			Set Items = Nothing
			
End Function

'Controls with a SetFromCache property  Only!
Public Function SetListItemsCollectionFromCache(obj,CacheName)
	obj.SetFromCache Application(CacheName)	
End Function

%>