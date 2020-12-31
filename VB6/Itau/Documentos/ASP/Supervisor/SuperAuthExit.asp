<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim RedirectForm

Main
%>
<%
Sub Main

	RedirectForm = "Supervisor.asp"

' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next

	Session("instructionElement").unlock False, False
		Session("errordebug") = "objVWInstructionElement.unlock"
		Call DisplayError

	Session("objRec").Close
	set Session("objRec") = nothing
	set Session("instructionElement") = nothing

	response.redirect RedirectForm

End Sub

%>

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<title></title>
</head>

<body>

<p>&nbsp;</p>

</body>
</html>
<%Response.Flush%>
