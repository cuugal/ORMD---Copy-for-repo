<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"--><%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Online Risk Register - Delete a Risk Assessment</title>
</head>
<body>
<%
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />") 
Escape = strReturn
End Function


dim val 
'dim conn1
dim conn
dim strSQL
dim rsDeleteQORA

val = cint(request.form("hdnQORAId"))
'response.write(val)

  set conn = Server.CreateObject("ADODB.Connection")
  'conn.open constr
  
  strSQL ="delete from tblQORA where numQORAId ="& val 
  
  set rsDeleteQORA = Server.CreateObject("ADODB.Recordset")
 ' rsDeleteQORA.Open strSQL, conn, 3, 3 
  
  conn.open constr
  conn.BeginTrans
  conn.Execute strSQL
  conn.commitTrans
  
  strSQL ="delete from tblRiskControls where numQORAID ="& val 
  
  set rsDeleteQORA = Server.CreateObject("ADODB.Recordset")
 ' rsDeleteQORA.Open strSQL, conn, 3, 3 
  
  'conn.open constr
  conn.BeginTrans
  conn.Execute strSQL
  conn.commitTrans
  
'else
%>

<!-- **************************************New Code that displays search results afer deleting -->


<!--#include file="reportAfterEdit.asp"-->
</body>
</html>
