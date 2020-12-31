<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim Supervisor, StartEmp, EndEmp

Dim user, password, router_URL

Dim indexName, queryFlags

Dim minValues(1), maxValues(1), queueNames(0)

Dim objVWQueue, objVWQueueQuery, objVWSession
Dim VWQueueElements

Dim OperationName()
Dim Employee()

Dim recovery

Dim RedirectForm, i

Main
%>
<%
Sub Main
	
' To trap errors un-remark following line.  Leave the line remarked for debugging
	On Error Resume Next

	RedirectForm = "SuperResult.asp"

	If (Request.Form("supervisor") = "") Then
		Session("processerror") = "Please select a Supervisor before proceeding"
		response.redirect RedirectForm
	End If

	If (Request.Form("startemp") = "") Then
		Session("processerror") = "Please enter a starting Employee value before proceeding"
		response.redirect RedirectForm
	End If

	If (Request.Form("endemp") = "") Then
		Session("processerror") = "Please enter an ending Employee value before proceeding"
		response.redirect RedirectForm
	End If

'Read the fields posted from the form
Supervisor = Request.Form("supervisor")
StartEmp = Request.Form("startemp")
EndEmp = Request.Form("endemp")

set objVWSession = Session("vwSession")

' 1) Recover User for the Supervisor queue
' 2) Create a queue object for the Supervisor queue
' 3) Start a query and set the following parameters
'
'      indexName
'      minValues
'      maxValues
'      queryFlags
'	   filter
'	   substitutionVars
'
'		Use the Supervisor index defined as Supervisor + Employee in Composer
'		Set the filter and substitutionVars to the value Empty
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
' 4) Fetch 50 queue elements
' 5) Put the operation name and employee name in the arrays
'		OperationName()
'		Employee()
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

	queueNames(0) = "Supervisor"
	recovery = objVWSession.recoverUser(Empty, queueNames)
		Session("errordebug") = "objVWSession.recoverUser"
		Call DisplayError

	set objVWQueue = objVWSession.getQueue("Supervisor")
		Session("errordebug") = "objVWSession.getQueue"
		Call DisplayError

	indexName = "Supervisor"
	queryFlags = 96
	minValues(0) = Supervisor
	minValues(1) = StartEmp
	maxValues(0) = Supervisor
	maxValues(1) = EndEmp

	set objVWQueueQuery = objVWQueue.startQuery(indexName, minValues, maxValues, queryFlags, Empty, Empty)
		Session("errordebug") = "objVWQueue.startQuery"
		Call DisplayError

	VWQueueElements = objVWQueueQuery.fetchQueueElements(50)
		Session("errordebug") = "objVWQueueQuery.fetchQueueElements"
		Call DisplayError

	if UBound(VWQueueElements) = -1 then
		Session("processerror") = "There are no items that match your specifications"
		response.redirect RedirectForm
	end if


	ReDim OperationName(UBound(VWQueueElements))
	ReDim Employee(UBound(VWQueueElements))

	for i = 0 to UBound(VWQueueElements)
		OperationName(i) = VWQueueElements(i).getOperationName()
			Session("errordebug") = "objVWQueueElements(i).getOperationName"
			Call DisplayError
		Employee(i) = VWQueueElements(i).getFieldValue("swcEmployee")
			Session("errordebug") = "objVWQueueElements(i).getFieldValue"
			Call DisplayError

	next

' ^^^^^^^^^^^^^^ END - INSERT CODE HERE - END ^^^^^^^^^^^^^^^^^^

' Save the Queue Elements in a Session variable for use on the
' next page.

	Session("queueElements") = VWQueueElements

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

	<form action="SuperProcess.asp" method="POST" name="SuperProcess" target="_top">
	<table border="0">
	<!-- Row 1 -->
	<tr>
		<!-- Col 1 -->
		<td align="left" colspan="2"><strong>Select the item to process</strong><br>
		</td>
	</tr>				
	<!-- Row 2 -->
	<tr>
		<!-- Col 1 -->
		<td><select name="intray" size="10" tabindex="1">
<%			for i = 0 to UBound(VWQueueElements)%>
				<option value="<%=i%>"> <%=OperationName(i)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=Employee(i)%></option>
<%			
			next
%>
				<option>
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
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					  &nbsp;&nbsp;&nbsp;
				  </option>
			</select>
		</td>
		<!-- Col 2 -->
		<td align="right" valign="top">
			<input type="submit" value="Process" tabindex="2">
		</td>
	</tr>
	<!-- Row 3 -->
	<tr>
		<td></td>
		<!-- Col 2 -->
		<td align="left" valign="top"><a href="../../../default.asp" target="_top">
			<img src="../../../Images/exit.gif" alt="Exit" border="0" WIDTH="79" HEIGHT="26"></a>
		</td>
	</table>
	</table>
	</form>
</body>
</html>
