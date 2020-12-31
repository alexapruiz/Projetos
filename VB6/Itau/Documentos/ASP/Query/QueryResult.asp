<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

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
		<td align="left"><strong>Work Objects found</strong><br>
		</td>
	</tr>				
	<!-- Row 2 -->
	<tr>
		<!-- 1 -->
		<td><select name="listbox" size="10" disabled>
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
			<input TYPE="text" name="txtnewname" size="60" maxlength="60" disabled>
		</td>
		<!-- 2 -->
		<td align="left" valign="top">
			<input type="submit" value="Change Name" disabled>
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
		alert('All Work Objects for Travel Request ' + requestid + processresult );

	if (processerror != "")
		alert( processerror )

	if (errorcode != 0)
		alert('This error occured during call ' + errordebug
			  + '\n\nThe Error Message is: \n' + errordesc);
-->
</script>


<%
Session("requestid") = 0
Session("processerror") = ""
Session("processresult") = ""
Session("errorcode") = 0
Session("errordesc") = "no error"
Session("errordebug") = "no error"
 %>

</body>
</html>
