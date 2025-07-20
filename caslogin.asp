<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="settings.asp"-->

<%

CAS_Server = "https://secure.its.yale.edu/cas/servlet/"
strCasPath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")


'See if already logged on
uid = Session("NetID")

If uid= "" then
	'Check for ticket returned by CAS redirect
	ticket = Request.QueryString("ticket")
	If ticket="" then
		'No session, no ticket, Redirect to CAS Logon page
		url = CAS_Server+"login?service="+strCasPath
		Response.Redirect(url)
	Else
		'Back from CAS, validate ticket and get userid
		'Response.Write "back from cas"
		Set theXmlObject = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		url =CAS_Server+"validate?ticket="+ticket+"&"+"service="+strCasPath

		theXmlObject.open "GET",url,false

		theXmlObject.send

		strRetVal = theXmlObject.responseText

		'GET LENGTH OF STRING AND LOOK FOR YES RESPONSE
		strLength = Len(strRetVal)
		strApproved = InStr(strRetVal,"yes")

		'IF RESPONSE DOES NOT CONTAIN YES, REDIRECT TO FAIL PAGE
		If strApproved=0 then
   			Response.Redirect "fail.html"
		Else
  			strStartNetID = strApproved+4
  			strNetIDChars = strLength-strStartNetID
  			strNetID = Mid(strRetVal,strStartNetID,strNetIDChars)
		End If
		Session("NetID")=strNetID
  	End If
End If

Response.Redirect "makepage.asp"



%>