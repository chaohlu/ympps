<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="settings.asp"-->
<!--#include virtual="/adovbs.inc"-->

<%

strSection = Request.QueryString("section")

If strSection = "ChromePC" then

%>

<p style="padding-left: 40px; font-weight: bold;">To view the source code of a website on a PC using Chrome, navigate to the page you want and use the keyboard shortcut: Control+U. You can also right-click on the page and select “View Page Source” from the dropdown menu.</p>

<%

End If

If strSection = "FirefoxPC" then

%>

<p style="padding-left: 40px; font-weight: bold;">To view the source code of a website on a PC using Firefox, navigate to the page you want and select Tools > Web Developer > Page Source. You can also right-click on the page and select “View Page Source” from the dropdown menu.</p>

<%

End If

If strSection = "EdgePC" then

%>

<p style="padding-left: 40px; font-weight: bold;">To view the source code of a website on a PC using Microsoft Edge, right-click on the page and select “View Page Source” from the dropdown menu.</p>

<%

End If



%>