<html><head>
<TITLE>formCheckBoxRespond.asp</TITLE>
</head><body bgcolor="#FFFFFF">
<%
If request.form("Correio")="on" then
    response.write "<br>N�s confirmaremos por Correio"
end if
If request.form("Sedex")="on" then
    response.write "<br>N�s confirmaremos por Sedex"
end if
If request.form("EMail")="on" then
    response.write "<br>N�s confirmaremos por EMail"
end if
If request.form("fax")="on" then
    response.write "<br>N�s confirmaremos por fax"
end if
If request.form("tel")="on" then
    response.write "<br>N�s confirmaremos por tel"
end if
%>
</body></html>
