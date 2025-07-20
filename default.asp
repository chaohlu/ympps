<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<%



%>

<!--#include virtual="/yppstemplate_d7/include_top_1.inc"-->


<link type="text/css" rel="stylesheet" href="styles.css" media="all">


<!--#include virtual="/yppstemplate_d7/include_top_2.inc"-->


 <!-- ++++++++++++++++++++++ PAGE TITLE ++++++++++++++++++++++++++ -->


Newsletter Web Page System


<!-- ++++++++++++++++++++++ END PAGE TITLE ++++++++++++++++++++++++++ -->    


<!--#include virtual="/yppstemplate_d7/include_bottom_1.inc"-->


 <!-- START CONTENT ---------------------------------------->  
    
    
<%
If Session("errormessage")<>"" then
	Response.Write "<p style='color:red;font-weight:bold;'>"& Session("errormessage") & "</p>"
End If
Session("errormessage") = ""
%>
<p><a href="caslogin.asp">Click here to log in</a></p>



 <!-- END CONTENT ------------------------------------------> 


<!--#include virtual="/yppstemplate_d7/include_bottom_2.inc"-->