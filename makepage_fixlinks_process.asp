<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%

strSubfolder = Request.Form("subfolder")
counterstart = Request.Form("counterstart")
counterend = Request.Form("counterend")
updateAction = Request.Form("action")

Dim DataConn, objRS, strSQL
Set DataConn = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
DataConn.Open strSQLConnect

'GET URL
strSQL = "SELECT * FROM tblEmailPagesWebversion WHERE Subfolder = ?"
Set cmd = Server.CreateObject("ADODB.COMMAND")
Set cmd.ActiveConnection = DataConn
cmd.Prepared = true
cmd.CommandType = adCmdText
cmd.CommandText = strSQL
cmd.Parameters.Append cmd.CreateParameter("Subfolder", adVarChar, adParamInput, 50, strSubfolder)
objRS.Open cmd,,1,3

If NOT objRS.EOF then
	strPageURL = objRS("URL")
	strNetID = objRS("NetID")
End If

objRS.Close
Set objRS = Nothing
DataConn.Close
Set DataConn = Nothing

'GET PAGE SOURCE CODE 
Dim httpRequest, postResponse
Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "GET", strPageURL, False
httpRequest.Send
pageSource = httpRequest.ResponseText

'*************************** START LINK REPLACE PROCESS *************************

pageSourceNew = pageSource

For i = counterstart To counterend

	radioName = "LinkOptions_" & i
	oldLinkField = "link_" & i & "_old"
	newLinkField = "link_" & i & "_new"
	linkItemField = "linkitem_" & i
	fullLinkField = "fulllinkcode_" & i
	radioResult = Request.Form(radioName)
	
	If radioResult = "ignore" OR radioResult = "" then
		'DO NOTHING
	ElseIf radioResult = "remove" then
		linkItem = Request.Form(linkItemField)
		fullLink = Request.Form(fullLinkField)
		pageSourceNew = Replace(pageSourceNew,fullLink,linkItem)
	ElseIf radioResult = "new" then
		oldLink = Request.Form(oldLinkField)
		newLink = Request.Form(newLinkField)
		pageSourceNew = Replace(pageSourceNew,oldLink,newLink)
	End If
	
	'response.write radioResult & "<BR>" & chr(10)

Next

If pageSourceNew <> pageSource then

	'UPDATE THE PAGE FILE
	strFolderPath = strFilePath & strNetID
	strSubFolderPath = strFolderPath & "\" & strSubfolder
	strFilePath = strSubFolderPath & "\default.html"

	Dim fso, MyFile, sFileName
	Set fso = CreateObject("Scripting.FileSystemObject")

	Set MyFile = fso.CreateTextFile(strFilePath, True, True)
	MyFile.WriteLine pageSourceNew

	'Response.Write pageSourceNew
	
End If

Response.Redirect "makepage_confirm.asp?action=" & updateAction & "&subfolder=" & strSubfolder

%>
