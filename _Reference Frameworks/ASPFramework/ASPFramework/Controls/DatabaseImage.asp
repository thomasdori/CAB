<%@ LANGUAGE="VBSCRIPT" %> 
<!--#Include File = "Samples/DBWrapper.asp" -->
<% 

Dim objRsImage  
Dim sSQL

'clear existing HTTP header 
Response.Expires = -10 
Response.Buffer = True 
Response.Clear 

'More on MIME types: http://www.utoronto.ca/ian/books/html4ed/appb/mimetype.html 
Response.ContentType = "image/jpeg" 
'SaveBinaryData server.mappath("\") & "\THEIMG", GetImageData(1)
  
sSQL = "SELECT BlobData FROM BlobTable Where BlobId = " & Replace(Request.QueryString("ID"),"'","''")
Set objRsImage = GetRecordSet(sSQL)
 
'use a BinaryWrite method call to output image to first page 
If objRsImage.RecordCount>0 Then
	Response.BinaryWrite objRsImage(0) 
End If
objRsImage.Close 
Set objRsImage = Nothing 
 
Response.End 

' CreateObject("Scripting.FileSystemObject")
'write image data To a disk
'This would be the main
'Dim ID
'ID = Request.QueryString("ID")
'If IsNumeric(ID) Then
'  Response.BinaryWrite GetImageData(ID)
'End If
'------>>> SaveBinaryData "C:\inetpub\root7\next.gif", GetImageData(2)

Function GetImageData(ID)
  Dim sSQL ,rs
  sSQL = "SELECT BlobData FROM BlobTable Where BlobId = " & Replace(Request.QueryString("ID"),"'","''")
  Set rs = GetRecordSet(sSQL)

  'GET binary data from recordset
  GetImageData = rs(0)  
  'Use this code instead of previous line For ORACLE.
'  GetImageData = RS("BinaryColumn").GetChunk( _
'      RS("BinaryColumn").ActualSize)
End Function


Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Function ReadBinaryFile(FileName)
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To get binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream
  BinaryStream.Open
  
  'Load the file data from disk To stream object
  BinaryStream.LoadFromFile FileName
  
  'Open the stream And get binary data from the object
  ReadBinaryFile = BinaryStream.Read
End Function
 


%>  

