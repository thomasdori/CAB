<!--#Include File = "Controls/StringBuilder.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<LINK rel="stylesheet" type="text/css" href="Samples.css">
</HEAD>
<BODY>
<form>
<input type="submit" name="HEY" value="Again"><br>
</form>
<%

	Class MyClass
		Dim Property1
		Public Output
		
		Private mMyProperty
		
		Private Sub Class_Initialize()
			Set Output = Response
		End Sub
		
		Public Property Get MyProperty
				MyProperty = mMyProperty
		End Property
		
		Public Property Let MyProperty(value)
			mMyProperty = value
		End Property
		
		Public Function CacheOutput()
			Set Output = new StringBuilder
			Output.GrowRate = 500
		End Function
	End Class	
		
	Dim varStart
	Dim x,tmp,obj
	Dim interval
	Set obj = new MyClass
	interval = 1000

	varStart = timer()
	For x = 0 To interval 
		tmp = obj.MyProperty
	Next	
	Response.Write "Get Public Property: " &  FormatNumber( Timer()-varStart,6) & "<BR>"

	varStart = timer()
	For x = 0 To interval 
		obj.MyProperty= "1"
	Next	
	Response.Write "Set Public Property: " &  FormatNumber( Timer()-varStart,6) & "<BR>"

	
	varStart = timer()
	For x = 0 To interval 
		tmp = obj.Property1
	Next	
	Response.Write "Get Public Field: " &  FormatNumber( Timer()-varStart,6) & "<BR>"

	varStart = timer()
	For x = 0 To interval 
		obj.Property1 = "1"
	Next	
	Response.Write "Set Public Field: " &  FormatNumber( Timer()-varStart,6) & "<BR>"


	varStart = timer()
	For x = 0 To interval 
		obj.Output.Write string(10,"A")
	Next	
	Response.Write "<BR>Total using Response: " &  FormatNumber( Timer()-varStart,6) & "<BR>"

	obj.CacheOutput
	varStart = timer()
	For x = 0 To interval 
		obj.Output.Append string(10,"A")
	Next	
	Response.Write obj.Output
	Response.Write "<BR>Total using StringBuilder: " &  FormatNumber( Timer()-varStart,6) & "<BR>"


%>


</BODY>
</HTML>
