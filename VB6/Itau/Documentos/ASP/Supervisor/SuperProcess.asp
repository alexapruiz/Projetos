<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%


Dim VWQueueElements
Dim objVWInstructionElement
Dim OperationName
Dim Selection

Dim Name, RequestNumber, ExpenseNumber
Dim RedirectForm

Main
%>
<%
Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next
	
	RedirectForm = "Supervisor.asp"

	If (Request.Form("intray") = "") Then
		Session("noselection") = 1
		response.redirect RedirectForm
	End If

	Selection = Request.Form("intray")

	VWQueueElements = Session("queueElements")

' 1) Fetch an Instruction Element for the selected item
'	 Lock the Instruction Element
'	 Do Not override lock
' 2) Get the Operation name for the Instruction Element
' 3) Perform a Select Case statement on the Operation Names
'	 Authorize
'	 ApproveExpenses
' 4) Place the Instruction Element in the Session variable "instructionElement"
' 5) Redirect to the page "SuperAuth.asp" or "SuperAppr.asp"
' 6) The operation ApproveExpenses has not been implemented in these exercises.
'
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	set objVWInstructionElement = VWQueueElements(Selection).fetchInstructionElement( true, false)
		Session("errordebug") = "VWQueueElements.fetchInstructionElement"
		Call DisplayError

	OperationName = objVWInstructionElement.getOperationName() 

	Select Case OperationName
		
		Case "Authorize"
			
			set Session("instructionElement") = objVWInstructionElement
			response.redirect "SuperAuth.asp"

		Case "ApproveExpenses"
			set Session("instructionElement") = objVWInstructionElement
			response.redirect "SuperAppr.asp"

		Case Else
			Session("processerror") = "Unknown operation.  Contact Programmer."
			response.redirect "SuperResult.asp"

	End Select

' vvvvvvvvvvvvvv END - INSERT CODE HERE - END vvvvvvvvvvvvvv 

Session("noselection") = 0
Session("errorcode") = 0
Session("errordesc") = "no error"
Session("errordebug") = "no error"

End Sub
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

</head>
<body TEXT="#2C396C" BGCOLOR="#D3D3D3">

</body>
</html>
