<%@ LANGUAGE=VBScript %>

<%	If ( Session("LoggedOn") = false ) Then
		Session("TargetPage") = "../Query/Query.asp"
		Session("TargetMessage") = "You must logon before performing a query."
		response.redirect "../Logon/logon.asp"
	End If
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">

<title>Travel Request Query</title>
</head>

<frameset rows="20%,*" border="0">
	<frame SRC="QueryForm.asp" name="Form" scrolling="NO">
	<frame SRC="QueryResult.asp" name="FormResults" scrolling="NO">
</frameset>

</html>
 