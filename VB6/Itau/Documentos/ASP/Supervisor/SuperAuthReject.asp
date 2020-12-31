<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim RedirectForm
Dim objRec

Main
%>
<%
Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next

	RedirectForm = "Supervisor.asp"

	set objRec = Session("objRec")
	objRec("Justified") = False
	objRec.Update
		Session("errordebug") = "Update 'Request'"
		Call DisplayError
	objRec.Close
	set objRec = Nothing
	set Session("objRec") = Nothing


	Session("instructionElement").setFieldValue "boprJustificationOK", False
		Session("errordebug") = "objVWInstructionElement.setFieldValue"
		Call DisplayError

	Session("instructionElement").unlock True, True
		Session("errordebug") = "objVWInstructionElement.unlock"
		Call DisplayError

	set Session("instructionElement") = nothing

	Session("processresult") = "rejected.\n\nThank you - notification sent."

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
