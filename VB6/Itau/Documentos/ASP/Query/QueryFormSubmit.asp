<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim txtQuery

Dim user, password, router_URL

Dim objVWSession, objVWRoster, objVWRosterQuery
Dim VWWorkObjects

Dim indexName, queryFlags

Dim minValues(0), maxValues(0)

Dim OperationName()
Dim WorkPerformerClassName()

Dim RedirectForm

Dim LoopCounter

Main
%>
<%
Sub Main
	
' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next

	RedirectForm = "QueryResult.asp"

	If (Request.Form("txtquery") = "") Then
		Session("processerror") = "Please enter the name of the Work Object you wish to search for."
		response.redirect RedirectForm
	End If

set objVWSession = Session("vwSession")

'Read the fields posted from the form
txtQuery = Request.Form("txtquery")

' 1) Get a Roster object for the CourseTravel Work Class
' 2) Start a roster query and set the following parameters:
'      indexName
'      minValues
'      maxValues
'      queryFlags
'      filter
'      substitutionVars
'
'    Set the filter and substitutionVars to the value Empty
'
'    The following system-defined index keys are defined for rosters: 
'      F_WobNum - Work Object number - type byte[] 
'      F_WobTag - The Work Object ID for the Work Class - type string
'
'		Query Flag Values
'		----------------------
'		NO OPTIONS = 0
'		READ LOCKED = 1
'		READ BOUND = 2
'		READ UNWRITABLE=4
'		LOCKED OBJECTS=16
'		MIN VALUE INCLUSIVE=32
'		MAX VALUE INCLUSIVE=64
'
'
' 3) Fetch 50 roster elements
' 4) Put the operation name and Work Performer Class name in the arrays
'		OperationName()
'		WorkPerformerClassName()
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	set objVWRoster = objVWSession.getRoster("CourseTravel")
		Session("errordebug") = "objVWSession.getRoster"
		Call DisplayError

	indexName = "F_WobTag"
	minValues(0) = txtQuery
	maxValues(0) = txtQuery
	queryFlags = 4+32+64

	set objVWRosterQuery = objVWRoster.startQuery(indexName, minValues, maxValues, queryFlags, Empty, Empty)
		Session("errordebug") = "objVWRoster.startQuery"
		Call DisplayError

	VWWorkObjects = objVWRosterQuery.fetchWorkObjects(50)
		Session("errordebug") = "objVWRosterQuery.fetchWorkObjects"
		Call DisplayError

	if UBound(VWWorkObjects) = -1 then
		Session("processerror") = "There are no items that match your specifications"
		response.redirect RedirectForm
	end if


	ReDim WorkPerformerClassName(UBound(VWWorkObjects))
	ReDim OperationName(UBound(VWWorkObjects))

	for LoopCounter = 0 to UBound(VWWorkObjects)
		WorkPerformerClassName(LoopCounter) = VWWorkObjects(LoopCounter).getWorkPerformerClassName()
			Session("errordebug") = "VWWorkObjects(LoopCounter).getWorkPerformerClassName"
			Call DisplayError
		OperationName(LoopCounter) = VWWorkObjects(LoopCounter).getOperationName()
			Session("errordebug") = "VWWorkObjects(LoopCounter).getOperationName"
			Call DisplayError

	next

' ^^^^^^^^^^^^^^ END - INSERT CODE HERE - END ^^^^^^^^^^^^^^^^^^

	Session("workObjects") = VWWorkObjects

	Session("processerror") = ""
	Session("processresult") = ""
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

	<form action="QueryProcess.asp" method="POST" name="QueryProcess" target="_top">
	<table border="0">
	<!-- Row 1 -->
	<tr>
		<!-- 1 -->
		<td align="left"><strong>Work Objects found</strong>
		</td>
	</tr>				
	<!-- Row 2 -->
	<tr>
		<!-- Col 1 -->
		<td><select name="listbox" size="10">
<%			for LoopCounter = 0 to UBound(VWWorkObjects)%>
				<option><%=WorkPerformerClassName(LoopCounter)%>&nbsp;&nbsp;&nbsp;<%=OperationName(LoopCounter)%>
<%			
			next
%>
				<option>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				  </option>
			</select>
		</td>
	</tr>
	<!-- Row 3 -->
	<tr>
		<!-- 1 -->
		<td align="left"><strong>New Employee Name</strong><br>
		</td>
	</tr>				


	<!-- Row 4 -->
	<tr>
		<!-- 1 -->
		<td>
			<input TYPE="text" name="txtnewname" size="60" maxlength="60" tabindex="1" onBlur="txtnewname.value=txtnewname.value.toUpperCase()">
		</td>
		<!-- 2 -->
		<td align="left" valign="top">
			<input type="submit" value="Change Name" tabindex="2">
		</td>
	<!-- Row 5 -->
	<tr>
		<!-- 1 -->
		<td>&nbsp;</td>
		<!-- 2 -->
		<td align="left" valign="top"><a href="../../../default.asp" target="_top">
			<img src="../../../Images/exit.gif" alt="Exit" border="0" WIDTH="79" HEIGHT="26"></a>
		</td>
	</table>
	</form>
</body>
</html>
