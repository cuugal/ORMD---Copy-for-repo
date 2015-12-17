<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Save SWMS</title>
     <!--#include file="bootstrap.inc"--> 
</head>
<%

'******************** code to fetch the values from the Create QORA Form **************************************

testval = request.form("hdnQORAID")
//Job Steps
strT4 = Request.form("T4")
temp = instr(1,strT4,"'",vbTextCompare)
      if temp <> 0 then 
         strT4 = Replace(strT4,"'","''",1)
      end if

'*************************Database connectivity Code***********************************************************

Dim conn
Dim rsAdd
Dim conn2
    Dim dte
    dte = Date()

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
  ' setting up the recordset
'***************************Insert into database**************************************************************
   ' DLJ 14April2010 added set boolSWMSRequired to catch saved SWMS even though not originaly selected as requiring one
   strSql = "Update tblQORA set strJobSteps = '"&strT4&"', boolSWMSRequired = true, dtDateCreated = '"&dte&"' where numQORAId = "&testval
   set rsAdd = Server.CreateObject("ADODB.Recordset")
  rsAdd.Open strSQL, conn, 3, 3
  
  
%>
      <!--#INCLUDE FILE="UpdateReview.asp"-->
<body>
     <!--#include file="HeaderMenu.asp"--> 
<div id="wrapper">
  <div id="content">
    <h1 class="pagetitle">SWMS <%=testval%> has been updated successfully</h1>
	
	  </div>
</div>

<!-- **************************************New Code that displays search results afer adding -->

<!--#include file="reportAfterEdit.asp"--></body>

</body>

</html>
