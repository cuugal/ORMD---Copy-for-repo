<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>
<%
dim loginId
loginId = session("strLoginId")
'Response.Write(loginId)
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Risk Level Report for Supervisors</title>
</head>
<%'*********************declaring the variables************************

  dim rsSearchH
  dim rsSearchM
  dim rsSearchL 
  dim rsFillFaculty
  dim rsFillLocation
  dim rsSearchFaculty
  dim Conn 
  dim strSQL
  dim strFacultyName
  dim strGivenName
  dim strSurname
  dim strName
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************writting the SQL ******************************
      
  '------------------------get the faculty for the login ---------------
  strSQL = "Select * from tblfacilitySupervisor,tblFaculty where tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId and tblFacilitySupervisor.strLoginId = '"& loginId &"'"
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  'Response.Write(strSQL) 
  rsSearchFaculty.Open strSQL, Conn, 3, 3     
  strFacultyName = rsSearchFaculty(7)     
  strGivenName = rsSearchFaculty(3)
  strSurname = rsSearchFaculty(4)
  strName = cstr(strGivenName) + " " + cstr(strSurname)
  Response.Write (strName)%><BR><%
  Response.Write(strFacultyName)
   
 '*********************Writting the report**************************** 
 ' code for sorting the risk levels = H
  strSQL = "SELECT * FROM tblQORA, tblFacility WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and strSupervisor = '"& loginId &"' and strAssessRisk = 'H' ORDER BY strRoomName"
  set rsSearchH = server.CreateObject("ADODB.Recordset")
      rsSearchH.Open strSQL, Conn, 3, 3  
  'Response.Write(strSQL)    
 while not rsSearchH.EOF 
    Response.Write(rsSearchH(18))%><BR><%
    rsSearchH.MoveNext  
 wend 
 '-----------------------------------------------------------------------
 ' code for sorting the risk levels = M
  strSQL = "SELECT * FROM tblQORA, tblFacility WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and strSupervisor = '"& loginId &"' and strAssessRisk = 'L' ORDER BY strRoomName"
  set rsSearchM = server.CreateObject("ADODB.Recordset")
      rsSearchM.Open strSQL, Conn, 3, 3 
  Response.Write("L") %><BR><% 
    
 while not rsSearchM.EOF 
    Response.Write(rsSearchM(18))%><BR><%
    rsSearchM.MoveNext  
 wend 
 
%>
<body>

</body>

</html>
