<%@ LANGUAGE=VBScript%>
<% Option Explicit %>
<%Response.Buffer=True%>
<%Response.Expires=0%>

<!-- #include file="..\..\..\Include\DisplayError.asp" -->

<%

Dim RedirectForm
Dim Logon
Dim objJig, objVWSession
Dim LogonName, LogonPassword, RouterName

Main

Sub Main

' To trap errors un-remark following line.  Leave the line remarked for debugging
'	On Error Resume Next

	RedirectForm = "logon.asp"

	Logon=Request.QueryString("Logon")

	If Session("LoggedOn") = true Then
		Session("vwSession").logoff
			Session("errordebug") = "objVWSession.logon"
			Call DisplayError
		Session("LoggedOn") = false
	End If

	If Logon = 0 Then

			set Session("vwSession") = nothing
			set Session("jig") = nothing
			
			response.redirect "../../../default.asp"
	End If		

	If (Request.Form("username") = "") and (Logon = 1) Then
		Session("processerror") = "Enter your User Name to logon"
		response.redirect RedirectForm
	End If

	If (Request.Form("router") = "") and (Logon = 1) Then
		Session("processerror") = "Enter a router URL. Ex. rmi://servername/routername"
		response.redirect RedirectForm
	End If

	LogonName = Request.Form("username")
	LogonPassword = Request.Form("password")
	RouterName = Request.Form("router")

'1.	Create the Java Interface object
'2. Create a VW Session object and save the	object vwSession for later use.  
'3. Use the LogonName, LogonPassword and RouterName from the Logon Form.
'4.	Log on to the VW Router.
'5. Check for errors using the DisplayError subroutine
' vvvvvvvvvvvvvv BEGIN - MODIFY CODE HERE - BEGIN vvvvvvvvvvvvvv 

	set objJig = Server.CreateObject("JiGlue.Util")

	set objVWSession = objJig.newInstance("filenet.vw.api.VWSession")

	objVWSession.logon LogonName, LogonPassword, RouterName
		Session("errordebug") = "objVWSession.logon"
		Call DisplayError

' ^^^^^^^^^^^^^^ END - INSERT CODE HERE - END ^^^^^^^^^^^^^^^^^^

	Session("LoggedOn") = true
	Session("LogonName") = LogonName
	Session("LogonPassword") = LogonPassword
	Session("RouterName") = RouterName
	set Session("vwSession") = objVWSession

	response.redirect Session("TargetPage")
	

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