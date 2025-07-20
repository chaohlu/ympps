<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%

strSubfolder = Request.QueryString("subfolder")
strAction = Request.QueryString("action")

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

%>


<!--#include virtual="/yppstemplate_d7/include_top_1.inc"-->

<link type="text/css" rel="stylesheet" href="styles.css" media="all">

<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>

<script language="javascript">

(function ($) {
  $(document).ready(function(){

    $(".linkcell a").on('click', function(event) {
            event.preventDefault();
            var url = $(this).attr('href');
            window.open(url, '_blank');
    });
    
    });
})(jQuery);

</script>


<!--#include virtual="/yppstemplate_d7/include_top_2.inc"-->


 <!-- ++++++++++++++++++++++ PAGE TITLE ++++++++++++++++++++++++++ -->


Update Links


<!-- ++++++++++++++++++++++ END PAGE TITLE ++++++++++++++++++++++++++ -->    


<!--#include virtual="/yppstemplate_d7/include_bottom_1.inc"-->


 <!-- START CONTENT ---------------------------------------->   
 
 <%
 If strAction = "update" then
	Response.Write "<p><a href='makepage.asp'>Return to main page</a></p>" & chr(10)
 End If
 %>
 
 <p>Your web page contains several links, which are listed below. If the original page has expired or the website has changed, and you need to replace a link embedded in your message or newsletter, please select from one of the options below.</p>
<p>You have three options for handling each of these links when the new web page is created:</p>
<ul>
  <li>You can LEAVE the link as-is. (This is the default option.)</li>
  <li>You can REMOVE the link from the web page. (If you choose this option, the text of the link will still appear on the page, but it will no longer be a clickable link.</li>
  <li>You can replace the link with a NEW direct link. If you choose this option, you should enter or paste the new URL in the text box.</li>
</ul>
<p><strong>PLEASE NOTE:</strong> Web URLs should be absolute, i.e. they should start with http: or https: -- or you can enter an email URL starting with "mailto:", e.g. "mailto:john.doe@yale.edu".</p>
 
<form action="makepage_fixlinks_process.asp" method="post" name="linksform" id="linksform">

<table cellspacing="0" cellpadding="10" border="1">
     
<%

'*************************** START LINK REPLACE PROCESS *************************

linkCounter = 0
counterStart = 1

arrPieces = Split(pageSource,"<a ")
for each x in arrPieces 
	If InStr(x,"href") > 0 then
		linkCounter = linkCounter + 1
		arrTemp = Split(x,"href=""")
		strLinkPart = arrTemp(1)
		arrLinkPieces = Split(strLinkPart,"""")
		strLink = arrLinkPieces(0)
		arrClickPieces = Split(strLinkPart,"</a>")
		arrClickSubPieces = Split(arrClickPieces(0),">")
		
		LinkItem = arrClickSubPieces(1)
		If InStr(LinkItem,"img") > 0 then
			LinkItem = LinkItem & ">"
		End If
		
		fullLinkContent = "<a " & arrTemp(0) & "href=""" & arrClickSubPieces(0) & ">" & LinkItem & "</a>"
		
		'Response.Write "Link " & linkCounter & ": " & strLink & "<BR>" & chr(10)
		'Response.Write "Click text " & linkCounter & ": " & LinkItem & "<BR><br>" & chr(10)
		
		'SHOW LINK INFO ONLY IF LINK IS NOT THE 'CLICK TO VIEW WEB PAGE' LINK
		If linkCounter = 1 AND LinkItem = "here." then
			counterStart = 2
			'DO NOTHING
		Else
			If InStr(strLink,"message.yale.edu") > 0 then
				Response.Write "<tr class='messagelink'>" & chr(10)
			Else
				Response.Write "<tr>" & chr(10)
			End If
			Response.Write "<td valign='top' class='linkcell' style='vertical-align: top;'>LINK CONTENT:<br />" & fullLinkContent & "</td>" & chr(10)
			Response.Write "<td valign='top' class='dont-break-out' style='vertical-align: top;'>LINK:<br />" & strLink & chr(10)
			Response.Write "<input type='hidden' id='link_" & linkCounter & "_old' name='link_" & linkCounter & "_old' value='" & strLink & "'>" & chr(10)
			Response.Write "<input type='hidden' id='linkitem_" & linkCounter & "' name='linkitem_" & linkCounter & "' value='" & LinkItem & "'>" & chr(10)
			Response.Write "<input type='hidden' id='fulllinkcode_" & linkCounter & "' name='fulllinkcode_" & linkCounter & "' value='" & fullLinkContent & "'>" & chr(10)
			Response.Write "</td>" & chr(10)
			Response.Write "<td>" & chr(10)
			Response.Write "<label><input type='radio' name='LinkOptions_" & linkCounter & "' value='ignore' checked='checked' id='LinkOptions_" & linkCounter & "_ignore'> Leave link as-is</label>" & chr(10)
			Response.Write "<label><input type='radio' name='LinkOptions_" & linkCounter & "' value='remove' id='LinkOptions_" & linkCounter & "_remove'> Remove link</label>" & chr(10)
			Response.Write "<label><input type='radio' name='LinkOptions_" & linkCounter & "' value='new' id='LinkOptions_" & linkCounter & "_new'> New link: </label>" & chr(10)
			Response.Write "<input type='text' name='link_" & linkCounter & "_new' id='link_" & linkCounter & "_new' size='40'>" & chr(10)
			Response.Write "</td>" & chr(10)
			Response.Write "</tr>" & chr(10)
		End If
		
	End If
next

%>

</table>

<input type="hidden" name="subfolder" id="subfolder" value="<%=strSubfolder%>">
<input type="hidden" name="counterstart" id="counterstart" value="<%=counterStart%>">
<input type="hidden" name="counterend" id="counterend" value="<%=linkCounter%>">
<input type="hidden" name="action" id="action" value="<%=strAction%>">

<p style="text-align: left;"><input type="submit" name="submit" id="submit" value="Submit"></p>

</form>

 <!-- END CONTENT ----------------------------------------- --> 


<!--#include virtual="/yppstemplate_d7/include_bottom_2.inc"-->
