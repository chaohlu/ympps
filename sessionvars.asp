<font face=arial size=1>
Session Variables - <% =Session.Contents.Count %> Found<br><br>
<%
Dim item, itemloop
For Each item in Session.Contents
If item<>"PagePhrases" AND item<>"ItemToShow" AND item<>"EstMultiQty" AND item<>"UserCart" AND item<>"rsBillToAddress" AND item<>"rsShipToAddress" AND item<>"rsMultiAddressList" AND item<>"ProductProfiles" AND item<>"vBillingFieldsValues" AND item<>"ProductDetails" AND item<>"CCArray" then
  If IsArray(Session(item)) then
    For itemloop = LBound(Session(item)) to UBound(Session(item))
%>
<% =item %>  <% =itemloop %> <font color=blue><% =Session(item)(itemloop) %></font><BR>
<%
    Next
  Else
%>
<% =item %> <font color=blue><% =Session.Contents(item) %></font><BR>
<%
  End If
End If
Next
%>


</font>
