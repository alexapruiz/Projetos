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
		<td><select name="intray" size="10" tabindex="1" disabled>
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
		<td align="left" valign="top">
			<input type="submit" value="Process" tabindex="2" disabled>
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
	</form>


<script LANGUAGE="JavaScript">
<!-- Display status and errors -->

<!--	
	var processresult = "<%=Session("processresult")%>";
	var processerror = "<%=Session("processerror")%>";
	var requestid = "<%=Session("requestid")%>";
	var requestername = "<%=Session("requestername")%>";
	var errorcode = <%=Session("errorcode")%>;
	var errordesc = "<%=Session("errordesc")%>";
	var errordebug = "<%=Session("errordebug")%>";



	if (processresult != "")
		alert('Travel Request ' + requestid +'  for ' + requestername + " " + processresult );

	if (processerror != "")
		alert( processerror );

	if (errorcode != 0)
		alert('This error occured during call ' + errordebug
			  + '\n\nThe Error Number is: \n' + errorcode
			  + '\n\nThe Error Message is: \n' + errordesc);
-->
</script>


<%
Session("processerror") = ""
Session("processresult") = ""
Session("requestid") = 0
Session("errorcode") = 0
Session("errordesc") = "no error"
Session("errordebug") = "no error"
 %>

</body>
</html>
