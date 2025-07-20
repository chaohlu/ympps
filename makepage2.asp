<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%
If Session("NetID")="" then
	Session("errormessage") = "You are not logged in."
	Response.Redirect "default.asp"
End If

Dim DataConn, objRS, strSQL
Set DataConn = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
DataConn.Open strSQLConnect

''' ADDED FOR DELETE FUNCTION
If Request.QueryString("action") = "delete" And IsNumeric(Request.QueryString("id")) Then
    Dim deleteCmd
    Set deleteCmd = Server.CreateObject("ADODB.Command")
    Set deleteCmd.ActiveConnection = DataConn
    deleteCmd.CommandType = adCmdText
    deleteCmd.CommandText = "DELETE FROM tblEmailPagesWebversion WHERE ID = ? AND NetID = ?"
    deleteCmd.Parameters.Append deleteCmd.CreateParameter("ID", adInteger, adParamInput, , CInt(Request.QueryString("id")))
    deleteCmd.Parameters.Append deleteCmd.CreateParameter("NetID", adVarChar, adParamInput, 10, Session("NetID"))
    deleteCmd.Execute
    Set deleteCmd = Nothing
    Response.Redirect "makepage.asp"
End If
''' END DELETE FUNCTION

%>

<!--#include virtual="/yppstemplate_d7/include_top_1.inc"-->

<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

<script type="text/javascript">
function dynamicContent(thePage,theDiv) {
	$(document).ready(function(){
		strPageToLoad = "dynamic_content.asp?section=" + thePage;
		$("#divone").html('');
		$("#divtwo").html('');
		$("#divthree").html('');
		$("#divfour").html('');
		$("#divfive").html('');
		$("#divsix").html('');
		$("#divseven").html('');
		$("#diveight").html('');
		thisContent = $(theDiv).html();
		if (thisContent=='') {
			$(theDiv).load(strPageToLoad);
		} else {
			$(theDiv).html('');
		}
	});
}
</script>

<link type="text/css" rel="stylesheet" href="styles.css" media="all">

<!--#include virtual="/yppstemplate_d7/include_top_2.inc"-->

<!-- ++++++++++++++++++++++ PAGE TITLE ++++++++++++++++++++++++++ -->
Make a Web Page from HTML Code
<!-- ++++++++++++++++++++++ END PAGE TITLE ++++++++++++++++++++++++++ -->    

<!--#include virtual="/yppstemplate_d7/include_bottom_1.inc"-->

<!-- START CONTENT ---------------------------------------->  

<p>This tool creates a web page from code that is created in the Yale Message tool. Here are the steps required to create a permanent link to your newsletter:</p>
<ul>
<li>Once you have created a message in the Yale Message system, navigate to the Content tab to select the Template-Based Email message that you created. You will need to select “Edit Content” of the message. Next select &gt;/&lt;Code View to see the HTML source code that was written to generate your newsletter. You will need to select all the code that is presented to you.</li>
<li>Select all the code that you see (click on the code and type Ctrl+A (PC) or Command+A (Mac)) and copy it (Ctrl+C (PC) or Command+C (Mac)).</li>
<li>Navigate back to the Make a Web Page from HTML Code page. Under the heading &lsquo;Paste the HTML source code below:$rsquo; paste your code using Ctrl + V if you are using a PC. Use Command + V if you are using a Mac. Click on the “Create Page” button.</li>
<li>The resulting page is the long-term web version of your email. All your original links within your newsletter will be retained.</li>
</ul>
<hr>
<h4 style="margin: 1em 0;">Use the form below to create your new page</h4>
<form id="form1" name="form1" method="post" action="makepage_process.asp">
<p><label for="pagetitle">Enter a reference title for your page (used for administrative purposes only):</label><input name="pagetitle" type="text" id="pagetitle" size="30"></p>
<p><label for="HTML_Source">Paste the HTML source code below:</label><textarea name="HTML_Source" id="HTML_Source" cols="60" rows="20"></textarea></p>
 <p><input type="submit" name="submit" id="submit" value="Create Page" /></p>
</form>

<%
Response.Write "<h4 style='margin: 4em 0 1em;'><hr />All web pages created by user " & Session("NetID") & "</h4>" & vbCrLf
Response.Write "<table id='allpages'><tbody>" & vbCrLf

Set cmd = Server.CreateObject("ADODB.COMMAND")
Set cmd.ActiveConnection = DataConn
cmd.Prepared = true
cmd.CommandType = adCmdText

strSQL = "SELECT * FROM tblEmailPagesWebversion WHERE NetID = ? ORDER BY Created DESC"
cmd.CommandText = strSQL
cmd.Parameters.Append cmd.CreateParameter("NetID", adVarChar, adParamInput, 10, Session("NetID"))
objRS.Open cmd,,1,3

While NOT objRS.EOF
	Response.Write "<tr><td>" & objRS("Title") & "</td>" & _
	    "<td>" & objRS("URL") & "</td>" & _
	    "<td>Created " & FormatDateTime(objRS("Created"),2) & "</td>" & _
	    "<td>" & _
	        "<a href='makepage_fixlinks.asp?action=update&subfolder=" & objRS("Subfolder") & "'>UPDATE LINKS</a> | " & _
	        "<a href='makepage.asp?action=delete&id=" & objRS("ID") & "' onclick='return confirm(""Are you sure you want to delete this page?"");'>DELETE</a>" & _
	    "</td>" & _
	    "<td><a href='" & objRS("URL") & "' target='_blank'>PREVIEW</a></td></tr>" & vbCrLf
	objRS.MoveNext
Wend

Response.Write "</tbody></table>" & vbCrLf

Set objRS = Nothing
DataConn.Close
Set DataConn = Nothing
%>

<!-- END CONTENT ------------------------------------------> 

<!--#include virtual="/yppstemplate_d7/include_bottom_2.inc"-->
