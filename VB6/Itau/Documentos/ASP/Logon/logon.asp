<%@ LANGUAGE=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Developer Studio">

<title>Visual WorkFlo Logon</title>
</head>
<body BGCOLOR="#D3D3D3" size="4">

<script LANGUAGE="JavaScript">
<!--
// SubmitAction is called when the Logon button is pressed.
// The form is submitted to LogonFormSubmit.asp, which is the Active Server Page that
// do the Visual WorkFlo logon/logoff
function SubmitAction()
{
document.Logon.submit();
}
//-->
</script>

		
		<% If Session("LoggedOn")=true Then %>
			<font COLOR="blue" size="4">You are currently logged on to</font>
			<font COLOR="black" size="4"><strong><%=Session("RouterName")%></strong></font>
		<% Else %>
			<font COLOR="blue" size="4"><%=Session("TargetMessage")%><br>
									Please logon to a Visual WorkFlo Router ...</font>
		<% End IF %> 

<form name="Logon" action="LogonFormSubmit.asp?Logon=1" method="POST">
<table BGCOLOR="#D3D3D3" TEXT="#000000" border="0 WIDTH=" 60%">

<!-- Row 1 -->
	<tr>
		<!--1--><td align="right"><strong>Username</strong></td>

		<!--2--><td colspan="2">
				<input type="text" maxlength="40" size="20" TABINDEX="1" name="username" value="<%=Session("LogonName")%>">
				
				</td>
		<!--3--><td align="center" valign="top" rowspan="2"><a href="javascript:SubmitAction()">
					<img src="../../../Images/logon.gif" alt="Logon" border="0" WIDTH="47" HEIGHT="43"><br>Logon</a>
				</td>
		<% If Session("LoggedOn")=true Then %>
		<!--4--><td>or</td>
		<!--5--><td align="center" valign="top" rowspan="2"><a href="LogonFormSubmit.asp?Logon=0" target="_self">
					<img src="../../../Images/logoff.gif" alt="Logoff" border="0" WIDTH="47" HEIGHT="43"><br>Logoff</a>
				</td>
		<% End If %>
	</tr>
<!-- Row 2 -->
	<tr>
		<!--1--><td align="right"><strong>Password</strong></td>

		<!--2--><td colspan="2">
				<input type="password" maxlength="40" size="20" TABINDEX="2" name="password" value="<%=Session("LogonPassword")%>">
				
				</td>
		<!--3-->
	</tr>

<!-- Row 3 -->
	<tr>
		<!--1--><td align="right"><strong>Router</strong></td>

		<!--2--><td colspan="2">
				<input type="text" maxlength="40" size="20" TABINDEX="3" name="router" value="<%=Session("RouterName")%>">
				
				</td>

		<!--3--><td align="left" valign="top" rowspan="2"><a href="../../../default.asp" target="_self">
														<img src="../../../Images/previous.gif" alt="Exit" border="0" WIDTH="45" HEIGHT="33"><br>Previous</a>
				</td>

	</tr>
</table>
</form>

<script LANGUAGE="JavaScript">
<!-- Display status and errors -->

<!--	
	var processresult = "<%=Session("processresult")%>";
	var processerror = "<%=Session("processerror")%>";
	var errorcode = <%=Session("errorcode")%>;
	var errordesc = "<%=Session("errordesc")%>";
	var errordebug = "<%=Session("errordebug")%>";

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
