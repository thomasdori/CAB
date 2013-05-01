<%

Class cASPCrypto
	'Get it from 
	'http://www.microsoft.com/downloads/details.aspx?FamilyID=860EE43A-A843-462F-ABB5-FF88EA5896F6&displaylang=en
	Private mSecret
	Private Sub Class_Initialize()
		mSecret = "XYZCLASPRULESXYZ"
	End Sub
		
	Public Function EncodeStr64(ByRef sInput)
		Dim oCrypto
		Set oCrypto = CreateObject("CAPICOM.EncryptedData")	
		oCrypto.SetSecret CStr(mSecret), 0
		oCrypto.Content = sInput
		EncodeStr64 = oCrypto.Encrypt
		Set oCrypto = Nothing
	End Function

	Public Function DecodeStr64(ByRef sEncoded)
		Dim oCrypto
		Set oCrypto = CreateObject("CAPICOM.EncryptedData")
		oCrypto.SetSecret mSecret, 0
		Call oCrypto.Decrypt(sEncoded)	
		DecodeStr64 = oCrypto.Content
		Set oCrypto = Nothing
	End Function
	
End Class

%>