<html><head>
<TITLE>ifrespond.asp</TITLE>
</head><body bgcolor="#FFFFFF">

<%

fname=request.querystring("FirstName")
lname=request.querystring("LastName")

response.write "First Name : " & request.querystring("FirstName")
response.write "Last  Name : " & request.querystring("LastName")

%>

</body></html>
