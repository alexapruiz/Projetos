<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim RedirectForm
Dim objVWInstructionElement, objRec, objJig

Main
%>
<%
Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next

	RedirectForm = "Supervisor.asp"


	set objVWInstructionElement = Session("instructionElement")
	set objRec = Session("objRec")
	objRec("Justified") = True
	objRec.Update
		Session("errordebug") = "Update 'Request'"
		Call DisplayError
	objRec.Close
	set objRec = Nothing
	set Session("objRec") = Nothing

' 1) Set the output operation parameter for the Instruction Element
'	 boprJustificationOK
' 2) Unlock the Instruction Element, save the values and dispatch
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	objVWInstructionElement.setFieldValue "boprJustificationOK", True
		Session("errordebug") = "objVWInstructionElement.setFieldValue"
		Call DisplayError

	objVWInstructionElement.unlock True, True
		Session("errordebug") = "objVWInstructionElement.unlock"
		Call DisplayError


' vvvvvvvvvvvvvv END - INSERT CODE HERE - END vvvvvvvvvvvvvv 

	Session("processresult") = "authorized.\n\nThank you - notification sent."
	set objVWInstructionElement = nothing
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
