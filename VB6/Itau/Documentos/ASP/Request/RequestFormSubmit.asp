<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->
<!-- #include file="..\..\..\Include\DataStore.inc" -->
<!-- #include file="..\..\..\Include\adovbs.inc" -->

<%

Dim Request_ID
Dim WorkClassName
Dim RedirectForm
Dim strConnect
Dim objJig, objVWSession, objVWRoster, objNewWorkObject, objRec

Main

Sub Main

	On Error Resume Next

	RedirectForm = "request.asp"

	If (Request.Form("requestername") = "") Then
		Session("processerror") = "Enter your Name to submit a travel request"
		response.redirect RedirectForm
	End If

	If (Request.Form("traveltype") = "business") Then
		Session("processerror") = " Business travel not implemented - change to Education"
		response.redirect RedirectForm
	End If

	WorkClassName = "CourseTravel"

	Call AddRequest
	Request_ID = CStr(objRec("Request"))

	set objVWSession = Session("vwSession")

' 1) Get a Roster object for the "CourseTravel" Work Class.
' 2) Create a new Work Object
' 3) Set the field value for swcRequest_ID using the Request_ID variable
' 4) Set the field values for the following from the Request.Form fields:
'		swcEmployee
'		swcCourseNumber
'		swcSupervisor_ID
'	Convert the variant form fields to the correct data types 
'	using the appropriate VBScript functions:
'		CStr() - Convert to String (VW type string)
'		CBool() - Convert to Boolean (VW type boolean)
'		CDbl() - Convert to Double (VW type float)
'		CLng() - Convert to Long (VW type integer)
' 5) Save the new object
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	set objVWRoster = objVWSession.getRoster(WorkClassName)
		Session("errordebug") = "objVWSession.getRoster"
		Call DisplayError

	set objNewWorkObject = objVWRoster.createWorkObject()
		Session("errordebug") = "objVWRoster.CreateWorkObject"
		Call DisplayError

	objNewWorkObject.setFieldValue "swcRequest_ID", Request_ID
		Session("errordebug") = "objNewWorkObject.setFieldValue swcRequest_ID"
		Call DisplayError

	objNewWorkObject.setFieldValue "swcEmployee", CStr(Request.Form("requestername"))
		Session("errordebug") = "objNewWorkObject.setFieldValue swcEmployee"
		Call DisplayError

	objNewWorkObject.setFieldValue "swcCourseNumber", CStr(Request.Form("course"))
		Session("errordebug") = "objNewWorkObject.setFieldValue swcCourseNumber"
		Call DisplayError

	objNewWorkObject.setFieldValue "swcSupervisor_ID", CStr(Request.Form("supervisor"))
		Session("errordebug") = "objNewWorkObject.setFieldValue swcSupervisor_ID"
		Call DisplayError

	objNewWorkObject.save
		Session("errordebug") = "objNewWorkObject.save"
		Call DisplayError

' ^^^^^^^^^^^^^^ END - INSERT CODE HERE - END ^^^^^^^^^^^^^^^^^^

	objRec.Close
	set objRec = nothing

	Session("requestid") = Chr(34) + Request_ID + Chr(34)
	response.redirect RedirectForm

End Sub

Sub AddRequest()

	set objRec = Server.CreateObject ("ADODB.Recordset")
		Session("errordebug") = "CreateObject ADODB.Recordset"
		Call DisplayError

	objRec.Open "Request", strConnect, adOpenStatic, adLockOptimistic, adCmdTable 
		Session("errordebug") = "Open 'Request'"
		Call DisplayError


	objRec.AddNew
		Session("errordebug") = "AddNew 'Request'"
		Call DisplayError

	objRec("Name") = Request.Form("requestername")
	objRec("DepartDate") = Request.Form("departdate")
	objRec("ReturnDate") = Request.Form("returndate")
	objRec("TravelReason") =Request.Form("travelreason")
	objRec("TravelType") = 	Request.Form("traveltype")
	objRec("Course") = 	Request.Form("course")
	objRec("CourseDesc") = 	Request.Form("course")
	objRec("Justified") = False
	objRec("Supervisor") = Request.Form("supervisor")

	objRec.Update
		Session("errordebug") = "Update 'Request'"
		Call DisplayError
		
	objRec.MoveLast
		Session("errordebug") = "MoveLast 'Request'"
		Call DisplayError
		

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