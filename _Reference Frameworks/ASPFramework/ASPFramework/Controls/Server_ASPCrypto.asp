<%

Class cASPCrypto
	'Get it from 
	'http://www.microsoft.com/downloads/details.aspx?FamilyID=860EE43A-A843-462F-ABB5-FF88EA5896F6&displaylang=en
	Public Function EncodeStr64(ByRef sInput)
		Dim oCrypto
		Set oCrypto = CreateObject("CAPICOM.Utilities")	
		EncodeStr64 = oCrypto.Base64Encode(sInput)	
		Set oCrypto = Nothing
	End Function

	Public Function DecodeStr64(ByRef sEncoded)
		Dim oCrypto
		Set oCrypto = CreateObject("CAPICOM.Utilities")
		
		DecodeStr64= oCrypto.Base64Decode(sEncoded)	
		Set oCrypto = Nothing
	End Function
	
End Class

%>