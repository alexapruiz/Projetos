<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->
<!-- #include file="..\..\..\Include\DataStore.inc" -->
<!-- #include file="..\..\..\Include\adovbs.inc" -->

<%

Dim user, password, router_URL

Dim VWWorkObjects, objWorkObject, txtNewName, txtRequestNumber
Dim VWExceptions

Dim strConnect, objRec

Dim RedirectForm, LoopCounter, LockError

Main
%>
<%
Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next
	
	RedirectForm = "Query.asp"

	If (Request.Form("txtnewname") = "") Then
		Session("processerror") = "Please enter the new name for the Work Object."
		response.redirect RedirectForm
	End If

	txtNewName = Request.Form("txtnewname")

	VWWorkObjects = Session("workObjects")

	txtRequestNumber =  VWWorkObjects(0).getFieldValue("swcRequest_ID")
			Session("errordebug") = "objWorkObject.getFieldValue"
			Call DisplayError

' 1) Lock the array of work objects - Do Not override locks
' 2) Check the VWException array for errors
' 3) If there are errors locking the work objects, unlock the objects you locked,
'	 set the error messages and exit the page.
' 4) Set the field value for swcEmployee to txtnewname for each object
' 5) Unlock the array of work objects saving the new value -  Do Not dispatch.
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	LockError = false
	VWExceptions = VWWorkObjects(0).lockMany (VWWorkObjects, false)
			Session("errordebug") = "VWWorkObjects(0).lockMany"
			Call DisplayError
	for LoopCounter = 0 to UBound(VWExceptions)
		if isObject(VWExceptions(LoopCounter)) Then
			LockError = true					
			Session("errordebug") = "VWWorkObjects(0).lockMany"
			Session("errorcode") = 1
			Session("errordesc") = VWExceptions(LoopCounter).getMessage()
			Exit For
		end if
	Next
	
	If LockError Then
		for LoopCounter = 0 to UBound(VWExceptions)
			if Not(isObject(VWExceptions(LoopCounter))) Then
				VWWorkObjects(LoopCounter).unlock false, false
			end if
		Next
		response.redirect RedirectForm
	End if
	
	for LoopCounter = 0 to UBound(VWWorkObjects)
		set objWorkObject = VWWorkObjects(LoopCounter)

		objWorkObject.setFieldValue "swcEmployee", txtNewName
			Session("errordebug") = "objWorkObject.setFieldValue"
			Call DisplayError

	next

	VWWorkObjects(0).unlockMany VWWorkObjects, True, False
		Session("errordebug") = "VWWorkObjects(0).unlockMany"
		Call DisplayError

' vvvvvvvvvvvvvv END - INSERT CODE HERE - END vvvvvvvvvvvvvv 

' Update the Travel Request Database
	Set objRec = Server.CreateObject ("ADODB.Recordset")
		Session("errordebug") = "CreateObject ADODB.Recordset"
		Call DisplayError

	objRec.Open "SELECT * FROM Request WHERE Request = " & txtRequestNumber & ";", _
				strConnect, adOpenForwardOnly, adLockOptimistic, adCmdText
		Session("errordebug") = "Open Select 'Request'"
		Call DisplayError

	objRec("Name") = txtNewName
	objRec.Update
		Session("errordebug") = "Update 'Request'"
		Call DisplayError
	objRec.Close

	Session("processresult") = " have been updated to the new name."
	Session("requestid") = txtRequestNumber
	
	response.redirect RedirectForm

End Sub
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

</head>
<body TEXT="#2C396C" BGCOLOR="#D3D3D3">

</body>
</html>
