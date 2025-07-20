<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%
Server.ScriptTimeout = 60 * 20

strSubfolder = Request.QueryString("subfolder")
updateAction = Request.QueryString("action")

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
End If

objRS.Close
Set objRS = Nothing
DataConn.Close
Set DataConn = Nothing

%>

<!--#include virtual="/yppstemplate_d7/include_top_1.inc"-->


<link type="text/css" rel="stylesheet" href="styles.css" media="all">


<!--#include virtual="/yppstemplate_d7/include_top_2.inc"-->


 <!-- ++++++++++++++++++++++ PAGE TITLE ++++++++++++++++++++++++++ -->

<%
If updateAction = "update" then
	Response.Write "Web Page Updated"
Else
	Response.Write "Web Page Created"
End If

%>


<!-- ++++++++++++++++++++++ END PAGE TITLE ++++++++++++++++++++++++++ -->    


<!--#include virtual="/yppstemplate_d7/include_bottom_1.inc"-->


 <!-- START CONTENT ---------------------------------------->  
    
     
<p>The URL for your <% If updateAction = "" then Response.Write "new " End If %>web page is</p>
<p><b style="color:blue;"><%=strPageURL%></b></p>
<p>Please copy this URL and use it wherever you need to link to a permanent web version of your email. You can preview the page by clicking <a href="<%=strPageURL%>" target="_blank">here</a>.
<%
If updateAction = "update" then
	Response.Write "<BR><em>PLEASE NOTE: if you have previewed this page previously, you may need to refresh the page to view any link updates.</em>"
End If
%>
</p>
<p><a href="makepage.asp">Return to main page</a></p>





 <!-- END CONTENT ------------------------------------------> 


<!--#include virtual="/yppstemplate_d7/include_bottom_2.inc"-->