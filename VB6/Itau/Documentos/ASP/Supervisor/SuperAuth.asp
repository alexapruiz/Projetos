<%@ LANGUAGE=VBScript %>
<% Option Explicit %>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->
<!-- #include file="..\..\..\Include\DataStore.inc" -->
<!-- #include file="..\..\..\Include\adovbs.inc" -->

<%
Dim objVWInstructionElement
Dim RequesterName, RequestNumber
Dim DepartDate, ReturnDate, TravelReason, Course
Dim strConnect
Dim objRec

Main
%>
<%
Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
  On Error Resume Next

set objVWInstructionElement = Session("instructionElement")


' 1) Get the operation parameters for the Instruction Element
'	 Put soprName in the variable RequesterName
'	 and soprRequestNumber in the variable RequestNumber
' vvvvvvvvvvvvvv BEGIN - INSERT CODE HERE - BEGIN vvvvvvvvvvvvvv 

RequesterName = objVWInstructionElement.getFieldValue("soprName")
	Session("errordebug") = "objVWInstructionElement.getFieldValue(soprName)"
	Call DisplayError

RequestNumber = objVWInstructionElement.getFieldValue("soprRequestNumber")
	Session("errordebug") = "objVWInstructionElement.getFieldValue(soprRequestNumber)"
	Call DisplayError

' vvvvvvvvvvvvvv END - INSERT CODE HERE - END vvvvvvvvvvvvvv 

Session("requestername") = RequesterName
Session("requestid") = RequestNumber

' Get the data from the Travel App Database

Set objRec = Server.CreateObject ("ADODB.Recordset")
	Session("errordebug") = "CreateObject ADODB.Recordset"
	Call DisplayError

objRec.Open "SELECT * FROM Request WHERE Request = " & RequestNumber & ";", _
			strConnect, adOpenForwardOnly, adLockOptimistic, adCmdText
	Session("errordebug") = "Open Select 'Request'"
	Call DisplayError

Set Session("objrec") = objRec

DepartDate = objRec("DepartDate")
ReturnDate = objRec("ReturnDate")
TravelReason = objRec("TravelReason")
Course = objRec("Course")

End Sub
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">

<title>Travel Authorization</title>
</head>
<body>

<form action="SuperAuthSubmit.asp" method="POST">

<table BGCOLOR="#D3D3D3" TEXT="#000000" border="0" WIDTH="60%">

<!-- Row 1 -->
	<tr>
		<!--1, 2 & 3--><td align="left" valign="bottom" colspan="3" width="25%"><strong>Please review the following travel request and then click on<br>authorize</strong></td>
		<!--4--><td width="25%" valign="center" rowspan="2"><img SRC="../../../Images/un-logo.gif" ALIGN="right" WIDTH="94" HEIGHT="73"></td>
	</tr>

<!-- Row 2 -->
	<tr>
		<!--1--><td width="25%" align="right"><strong>Request ID</strong></td>

		<!--2--><td width="25%" colspan="2">
				<input type="text" value="<%=RequestNumber%>" maxlength="20" size="47" disabled name="requestid">				
				</td>
		<!--3-->
		<!--4-->
	</tr>
<!-- Row 3 -->
	<tr>
		<!--1--><td width="25%" align="right"><strong>Name</strong></td>

		<!--2--><td width="25%" colspan="2">
				<input type="text" value="<%=RequesterName%>" maxlength="20" size="47" disabled name="requestername">
				</td>

		<!--3-->
		<!--4-->

	</tr>
<!-- Row 4 -->
		<!--1--><td width="25%"></td>
		<!--2--><td width="25%" align="center"><strong>Depart</strong></td>
		<!--3--><td width="25%" align="center"><strong>Return</strong></td>
		<!--4-->
<!-- Row 5 -->
	<tr>
		<!--1--><td width="25%" align="right" nowrap><strong>Travel Dates</strong></td>
		
		<!--2--><td width="25%" align="left">
				<input type="text" value="<%=DepartDate%>" size="21" disabled name="departdate">
				</td>

		<!--3--><td width="25%" align="right">
				<input type="text" value="<%=ReturnDate%>" size="21" disabled name="returndate">
				</td>

		<!--4-->
	</tr>


<!-- Row 6 -->
	<tr>
		<!--1--><td align="right" width="25%" valign="top"><strong>Travel<br>Justification</strong></td>

		<!--2--><td width="25%" colspan="2" rowspan="2"><textarea NAME="travelreason" ROWS="10" COLS="40" disabled><%=TravelReason%></textarea></td>
		<!--3-->
		<!--4--><td align="center" valign="center" width="25%">
				<input TYPE="submit" VALUE="AUTHORIZE" TABINDEX="1"></td>
	</tr>

<!-- Row 7 -->
	<tr>
		<!--1--><td width="25%">&nbsp;</td>
		<!--2-->
		<!--3-->
		<!--4--><td align="center" valign="top" width="25%"><a href="SuperAuthReject.asp" target="_self">
														<img src="../../../Images/reject.gif" alt="Reject" border="0" WIDTH="82" HEIGHT="27"></a>
				</td>
	</tr>
<!-- Row 8 -->
	<tr>
		<!--1--><td align="right" width="25%" nowrap><strong>Travel Purpose</strong></td>

		<!--2--><td><input TYPE="radio" NAME="travelpurpose" VALUE="education" CHECKED disabled><strong>Education</strong>
				</td>

		<!--3--><td><input TYPE="radio" NAME="travelpurpose" VALUE="business" disabled><strong>Business</strong>
				</td>
		<!--4--><td align="center" valign="top" width="25%"><a href="SuperAuthExit.asp" target="_self">
														<img src="../../../Images/exit.gif" alt="Exit" border="0" WIDTH="79" HEIGHT="26"></a>
				</td>
	</tr>
<!-- Row 9 -->
	<tr>
		<!--1--><td align="right" width="25%"><strong>Course</strong></td>
		<!--2--><td width="25%" align="left" colspan="2">
				<input type="text" value="<%=Course%>" size="47" disabled name="course">
				</td>
	</tr>
</table>
</form>

<!-- Display the Visual WorkFlo Error -->
<script LANGUAGE="JavaScript">
<!-- 'Display the Visual WorkFlo Error
	var errorcode = <%=Session("errorcode")%>;
	var errordesc = "<%=Session("errordesc")%>";
	var errordebug = "<%=Session("errordebug")%>";

	if (errorcode != 0)
		alert('This error occured during call ' + errordebug
			  + '\n\nThe Error Message is: \n' + errordesc);
-->
</script>


<%
Session("errorcode") = 0
Session("errordesc") = "no error"
Session("errordebug") = "no error"
 %>

</body>
</html>
