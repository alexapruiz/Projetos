<%@ LANGUAGE=VBScript %>

<%	If ( Session("LoggedOn") = false ) Then
		Session("TargetPage") = "../Supervisor/Supervisor.asp"
		Session("TargetMessage") = "You must logon before processing work as a Supervisor."
		response.redirect "../Logon/logon.asp"
	End If
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">

<title>Supervisor In-Tray</title>
</head>

<frameset rows="32%,*" border="0">
	<frame SRC="SuperForm.asp" name="Form" scrolling="NO">
	<frame SRC="SuperResult.asp" name="FormResults" scrolling="NO">
</frameset>

</html>
 