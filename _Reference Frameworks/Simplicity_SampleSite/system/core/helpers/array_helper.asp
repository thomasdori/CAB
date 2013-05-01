<%

' ---------------------------------------------------------------
'  Array Helper
' ---------------------------------------------------------------
' 
' add_attributes
' array_min
' array_max


' ---------------------------------------------------------------
'  Create Attributes from an Array
' ---------------------------------------------------------------
' 
'  @string

function add_attributes(var)
	
	dim attributes, atts
	
	if isarray(var) then
	
		attributes = var
		
		for i = 0 to ubound(attributes)
			
			atts = atts & attributes(i) & "=""" & attributes(i+1) & """ "
			
			i = i + 1
		
		next
	
	end if
	
	add_attributes =  atts
	
end function



' ---------------------------------------------------------------
'  Create Attributes from an Array
' ---------------------------------------------------------------
' 
'  @string

function add_pop_attributes(var)
	
	dim attributes, atts
	
	if isarray(var) then
	
		attributes = var
		
		for i = 0 to ubound(attributes)
			
			atts = atts & attributes(i) & "=" & attributes(i+1) & ","
			
			i = i + 1
		
		next
	
	end if
	
	atts = left(atts, len(atts)-1)
	
	add_pop_attributes =  atts
	
end function



' ---------------------------------------------------------------
'  Find the MIN value in an Array
' ---------------------------------------------------------------
' 
'  @array

function array_min(aNumberArray)
	
	dim dblLowestSoFar
	
	dblLowestSoFar = Null
	
	for I = LBound(aNumberArray) to UBound(aNumberArray)
		
		if IsNumeric(aNumberArray(I)) then
		
			if CDbl(aNumberArray(I)) < dblLowestSoFar or IsNull(dblLowestSoFar) then
			
				dblLowestSoFar = CDbl(aNumberArray(I))
			
			end if
		
		end if
	
	next
	
	array_min = dblLowestSoFar

end function



' ---------------------------------------------------------------
'  Find the MAX value in an Array
' ---------------------------------------------------------------
' 
'  @array

Function array_max(aNumberArray)
	
	dim dblHighestSoFar
	
	dblHighestSoFar = Null
	
	for I = LBound(aNumberArray) to UBound(aNumberArray)
		
		if IsNumeric(aNumberArray(I)) then
			
			if CDbl(aNumberArray(I)) > dblHighestSoFar or IsNull(dblHighestSoFar) then
				
				dblHighestSoFar = CDbl(aNumberArray(I))
			
			end if
		end if
	
	next	
	
	array_max = dblHighestSoFar

end function


' ---------------------------------------------------------------
'  Sort Array
' ---------------------------------------------------------------
' 
'  @array

Private Function ExchangeSort(byVal unsortedarray)
	Dim Front, Back, I, Temp, Arrsize, Excom, Exswap
	Arrsize = UBOUND( unsortedarray )
	For Front = 0 To Arrsize
		For Back = Front To Arrsize
			Excom = Excom + 1
			If unsortedarray(Front) > unsortedarray(Back) Then
				Temp = unsortedarray(Front)
				unsortedarray(Front) = unsortedarray(Back)
				unsortedarray(Back) = Temp
				Exswap = Exswap + 1
			End If
		Next
	Next
	ExchangeSort = unsortedarray
End Function


function array_check(array_vals, val)
	dim item
	for each item in array_vals
		
		if lcase( trim( item ) ) = lcase( trim( val ) ) then
			
			array_check = item
			
			exit for
			
		end if
	
	next
end function



function make_array(val, delimiter)
	
	make_array = Split(val, delimiter, -1, 1)
	
end function
%>