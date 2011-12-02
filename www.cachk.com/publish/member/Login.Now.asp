<%
Dim strSQL				' Structured Query Language
Dim objConn				' Database Connection
Dim objRs				' Recordset
Dim strLoginName		' Name user logs on with
Dim strLoginPassword	' Password to login with
Dim lngMemberID			' MemberID assigned to user account
Dim Marks

' Grab Form Data
strLoginName = Request.Form("LoginName")
strLoginPassword = Request.Form("LoginPassword")

' Open Database
Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")
objConn.Open "DRIVER=Microsoft Access Driver (*.mdb);uid=admin;pwd=z1x2c3v4b5n6m7;DBQ=" & Server.MapPath("Remote/Members.mdb")

' Look for User
'strSQL = "SELECT MemberID FROM Members WHERE LoginName = '" & strLoginName & "' AND LoginPassword = '" & strLoginPassword & "'"
strSQL = "SELECT * FROM Members WHERE LoginName = '" & strLoginName & "' AND LoginPassword = '" & strLoginPassword & "'"

Set objRs = objConn.Execute(strSQL)

' Notify visitor if account was found.
If objRs.EOF Then
	Response.Write("Login Failed")
Else
	lngMemberID = objRs(0)
	strLoginName = objRs(1)
	Marks = objRs(3)
	Response.Write("<p><font face='Arial' size=5>VIP Information</p></font>")
	Response.Write("<hr color=red>")
'	Response.Write("<p><font face='Arial' size=4>MemberID = " & lngMemberID & "</p></font>")
	Response.Write("<p><font face='Arial' size=4>VIP Name : " & strLoginName & "</p></font>")
	Response.Write("<p><font face='Arial' size=4>Marks : " & Marks & "</p></font>")
    Response.Write("<hr color=red>")
End If

' Garbage Collection
Set objRs = Nothing
objConn.Close
Set objConn = Nothing
%>