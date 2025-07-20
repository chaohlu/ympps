<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%
Server.ScriptTimeout = 60 * 20

If Session("NetID")="" then
	Session("errormessage") = "You are not logged in."
	Response.Redirect "default.asp"
End If


Dim DataConn, objRS, strSQL
Set DataConn = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
DataConn.Open strSQLConnect

strSQL="SELECT * FROM tblEmailPagesWebversion"
objRS.Open strSQL,DataConn,1,3

'MAKE URL FROM TIMESTAMP AND NETID
strQueryString = Now
strQueryString = Replace(strQueryString," AM","")
strQueryString = Replace(strQueryString," PM","")
strQueryString = Replace(strQueryString," ","")
strQueryString = Replace(strQueryString,"/","")
strQueryString = Replace(strQueryString,":","")

strFolderPath = strFilePath & Session("NetID")
strSubFolderPath = strFolderPath & "\" & strQueryString
strFilePath = strSubFolderPath & "\default.html"
strPageURL = strMakePagePath & Session("NetID") & "/" & strQueryString & "/"
strImageFolderPath = strSubFolderPath & "\images"
strDocFolderPath = strSubFolderPath & "\docs"

'CREATE FOLDERS AND FILE
Dim fso, MyFile, sFileName
Set fso = CreateObject("Scripting.FileSystemObject")
   
If NOT (fso.FolderExists(strFolderPath)) then 
    fso.CreateFolder(strFolderPath)
End If
If NOT (fso.FolderExists(strSubFolderPath)) then 
    fso.CreateFolder(strSubFolderPath)
End If
If NOT (fso.FolderExists(strImageFolderPath)) then 
    fso.CreateFolder(strImageFolderPath)
End If
If NOT (fso.FolderExists(strDocFolderPath)) then 
    fso.CreateFolder(strDocFolderPath)
End If

Set MyFile = fso.CreateTextFile(strFilePath, True, True)

'Response.Write strFilePath & "<BR>" & chr(10)

'*************************** START IMAGE MOVE/REPLACE PROCESS *************************
strPageTitle = Request.Form("pagetitle")
strSource = Request.Form("HTML_Source")
strFinalHTML = strSource

arrPieces = Split(strSource,"src=""")
for each x in arrPieces 
	If InStr(x,"http") = 1 AND ((InStr(x,".jpg") > 0) OR (InStr(x,".jpeg") > 0) OR (InStr(x,".png") > 0)) then
		'GET IMAGE URL STRING
		arrTemp = Split(x,"""")
		strURL = arrTemp(0)
		'GET IMAGE FILENAME
		arrURLPieces = Split(strURL,"/")
		strImageFile = arrURLPieces(UBound(arrURLPieces))
		'Response.Write strImageFile & "<BR>" & chr(10)

		strHDLocation = strImageFolderPath & "\" & strImageFile
		strImageFinalURL = strMakePagePath & Session("NetID") & "/" & strQueryString & "/images/" & strImageFile
		'Response.Write strHDLocation & "<BR>" & chr(10)
		'Response.Write strImageFinalURL & "<BR><BR>" & chr(10)

		' Fetch the file
		Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXMLHTTP.Open "GET", strURL, False
		objXMLHTTP.Send()

		If objXMLHTTP.Status = 200 Then
			Set objADOStream = CreateObject("ADODB.Stream")
			objADOStream.Open
			objADOStream.Type = 1 'adTypeBinary

			objADOStream.Write objXMLHTTP.ResponseBody
			objADOStream.Position = 0 'Set the stream position to the start

			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FileExists(strHDLocation) Then objFSO.DeleteFile strHDLocation
			Set objFSO = Nothing

			objADOStream.SaveToFile strHDLocation
			objADOStream.Close
			Set objADOStream = Nothing
			
			'REPLACE OLD IMAGE URL WITH NEW ONE IN HTML CODE
			strFinalHTML = Replace(strFinalHTML,strURL,strImageFinalURL)
			
		End if

		Set objXMLHTTP = Nothing
		
	End If
next


'*************************** START LINKED DOCUMENT MOVE/REPLACE PROCESS *************************

docCounter = 0
arrPieces2 = Split(strFinalHTML,"data-linkto=""document"" href=""")
for each x in arrPieces2 
	If InStr(x,"http") = 1 AND (InStr(x,"message.yale.edu") > 0) then
		'GET DOCUMENT URL STRING
		arrTemp = Split(x,"""")
		strURL = arrTemp(0)
		Response.Write strURL & "<BR>" & chr(10)
		'GET DOCUMENT FILENAME
		docCounter = docCounter + 1
		strDocFile = "document" & docCounter & ".pdf"

		strHDLocation = strDocFolderPath & "\" & strDocFile
		strDocFinalURL = strMakePagePath & Session("NetID") & "/" & strQueryString & "/docs/" & strDocFile
		Response.Write strHDLocation & "<BR>" & chr(10)
		Response.Write strDocFinalURL & "<BR><BR>" & chr(10)

		' Fetch the file
		Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXMLHTTP.Open "GET", strURL, False
		objXMLHTTP.Send()

		If objXMLHTTP.Status = 200 Then
			Set objADOStream = CreateObject("ADODB.Stream")
			objADOStream.Open
			objADOStream.Type = 1 'adTypeBinary

			objADOStream.Write objXMLHTTP.ResponseBody
			objADOStream.Position = 0 'Set the stream position to the start

			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FileExists(strHDLocation) Then objFSO.DeleteFile strHDLocation
			Set objFSO = Nothing

			objADOStream.SaveToFile strHDLocation
			objADOStream.Close
			Set objADOStream = Nothing
			
			'REPLACE OLD IMAGE URL WITH NEW ONE IN HTML CODE
			strFinalHTML = Replace(strFinalHTML,strURL,strDocFinalURL)
			
		End if

		Set objXMLHTTP = Nothing
		
	End If
next

'ADD JQUERY TO HIDE WEB LINKS
strFinalHTML = strFinalHTML & chr(10) & chr(10) & "<script src=""https://ajax.googleapis.com/ajax/libs/jquery/3.6.1/jquery.min.js""></script>" & chr(10)
strFinalHTML = strFinalHTML & "<script>" & chr(10)
strFinalHTML = strFinalHTML & "(function ($) {" & chr(10)
strFinalHTML = strFinalHTML & "$(document).ready(function(){" & chr(10)

strFinalHTML = strFinalHTML & "$( ""font:contains('To view this email as a web page')"" ).css( 'display', 'none' );" & chr(10)
strFinalHTML = strFinalHTML & "$( ""a:contains('View as Webpage')"" ).css( 'display', 'none' );" & chr(10)

strFinalHTML = strFinalHTML & "});" & chr(10)
strFinalHTML = strFinalHTML & "})(jQuery);" & chr(10)
strFinalHTML = strFinalHTML & "</script>" & chr(10)

MyFile.WriteLine strFinalHTML

Set fso = Nothing

 
objRS.AddNew
objRS("Title") = strPageTitle
objRS("URL") = strPageURL
objRS("NetID") = Session("NetID")
objRS("Subfolder") = strQueryString
objRS("Created") = Now
objRS.Update

objRS.Close
Set objRS = Nothing
DataConn.Close
Set DataConn = Nothing

Response.Redirect "makepage_confirm.asp?subfolder=" & strQueryString

'CHECK FOR LINKS AND REDIRECT
'If InStr(strFinalHTML,"<a ") > 0 Then
'	Response.Redirect "makepage_fixlinks.asp?subfolder=" & strQueryString
'Else
'	Response.Redirect "makepage_confirm.asp?subfolder=" & strQueryString
'End If


%>
