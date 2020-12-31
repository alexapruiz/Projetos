<%@ LANGUAGE=VBScript %>

<%	If ( Session("LoggedOn") = false ) Then
		Session("TargetPage") = "../Request/Request.asp"
		Session("TargetMessage") = "You must logon before creating a Travel Request."
		response.redirect "../Logon/logon.asp"
	End If
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">

<title>Travel Request</title>
</head>
<body>
 
<form action="RequestFormSubmit.asp" method="POST">

<table BGCOLOR="#D3D3D3" TEXT="#000000" border="0" WIDTH="60%">

<!-- Row 1 -->
	<tr>
		<!--1 & 2--><td align="right" valign="bottom" colspan="2" width="25%"><strong>Enter your travel information</strong></td>
		<!--3--><td width="25%">&nbsp;</td>
		<!--4--><td width="25%" valign="center" rowspan="2"><img SRC="../../../Images/un-logo.gif" ALIGN="right" WIDTH="94" HEIGHT="73"></td>
	</tr>

<!-- Row 2 -->
	<tr>
		<!--1--><td width="25%" align="right"><strong>Name</strong></td>

		<!--2--><td width="25%" colspan="2">
				<input type="text" maxlength="20" size="47" TABINDEX="1" name="requestername" onBlur="requestername.value=requestername.value.toUpperCase()">
				
				</td>
		<!--3-->
		<!--4-->
	</tr>
<!-- Row 3 -->
	<tr>
		<!--1--><td width="25%" align="right"><strong>Supervisor</strong></td>

		<!--2--><td width="25%" colspan="2">
				<select name="supervisor" wsize="1" TABINDEX="2"> 
                <option value="BRENDA TAYLOR" selected>BRENDA TAYLOR</option>
                <option value="CAROL BAKER">CAROL BAKER</option>
                <option value="HAROLD WEST">HAROLD WEST</option>
                <option value="JERRY FALCON">JERRY FALCON</option>
                <option value="JIM CONNOR">JIM CONNOR</option>
				<option value="LINDA ADAMS">LINDA ADAMS</option>
				<option value="MIKE GALLO">MIKE GALLO</option>
				<option value="PAULA STEVENS">PAULA STEVENS
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											  &nbsp;&nbsp;&nbsp;&nbsp;</option>
				</select></td>

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
			<select name="departdate" size="1" TABINDEX="3">
 			<%For i=1 to 30%>
  					<option value="<%=Date + i%>" <%If i = 1 Then%> selected <%End If%>> 
					<%=Date + i%>
					<%If i = 1 Then%>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;
					<%End If%> 					
					</option>
			<%Next%>
            </select></td>

		<!--3--><td width="25%" align="right">
			<select name="returndate" size="1" TABINDEX="4">
			<%For i=1 to 30%>
  					<option value="<%=Date + i%>" <%If i = 1 Then%> selected <%End If%>> 
					<%=Date + i%>
					<%If i = 1 Then%>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;
					<%End If%> 					
					</option>
			<%Next%>
            </select></td>

		<!--4-->
	</tr>


<!-- Row 6 -->
	<tr>
		<!--1--><td align="right" width="25%" valign="top"><strong>Travel<br>Justification</strong></td>

		<!--2--><td width="25%" colspan="2" rowspan="2"><textarea NAME="travelreason" ROWS="10" COLS="40" TABINDEX="5"></textarea></td>
		<!--3-->
		<!--4--><td align="center" valign="center" width="25%">
				<input TYPE="submit" VALUE="SUBMIT" TABINDEX="9"></td>
	</tr>

<!-- Row 7 -->
	<tr>
		<!--1--><td width="25%">&nbsp;</td>
		<!--2-->
		<!--3-->
		<!--4--><td align="center" valign="top" width="25%">
				<input TYPE="reset" VALUE="CANCEL" TABINDEX="10"></td>
	</tr>
<!-- Row 8 -->
	<tr>
		<!--1--><td align="right" width="25%" nowrap><strong>Travel Purpose</strong></td>

		<!--2--><td><input TYPE="radio" NAME="traveltype" VALUE="education" CHECKED TABINDEX="6"><strong>Education</strong>
				</td>

		<!--3--><td><input TYPE="radio" NAME="traveltype" VALUE="business" TABINDEX="7"><strong>Business</strong>
				</td>
		<!--4--><td align="center" valign="top" width="25%"><a href="../../../default.asp" target="_self">
														<img src="../../../Images/exit.gif" alt="Exit" border="0" WIDTH="79" HEIGHT="26"></a>
				</td>
	</tr>
<!-- Row 9 -->
	<tr>
		<!--1--><td align="right" width="25%"><strong>Course</strong></td>
		<!--2--><td width="25%" align="left" colspan="2">
			<select name="course" size="1" TABINDEX="8">
                <option value="200191 Visual WorkFlo Installation and Administration">200191 Visual WorkFlo Installation and Administration</option>
				<option value="200453 Visual WorkFlo Application Design" selected>200453 Visual WorkFlo Application Design</option>
                <option value="200454 Visual WorkFlo Application Development">200454 Visual WorkFlo Application Development</option>
            </select></td>
	</tr>
</table>
</form>

<script LANGUAGE="JavaScript">
<!-- Display status and errors -->

<!--	
	var processresult = "<%=Session("processresult")%>";
	var processerror = "<%=Session("processerror")%>";
	var requestid = <%=Session("requestid")%>;
	var requestername = "<%=Session("requestername")%>";
	var errorcode = <%=Session("errorcode")%>;
	var errordesc = "<%=Session("errordesc")%>";
	var errordebug = "<%=Session("errordebug")%>";

	if (requestid != 0)
		alert('Your request number ' + requestid +' has been sent for approval.');


	if (processerror != "")
		alert( processerror )

	if (errorcode != 0)
		alert('This error occured during call ' + errordebug
			  + '\n\nThe Error Message is: \n' + errordesc);
-->
</script>

<%
Session("processerror") = ""
Session("requestid") = 0
Session("errorcode") = 0
Session("errordesc") = "no error"
Session("errordebug") = "no error"
 %>

</body>
</html>
