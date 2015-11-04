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
  dim dtDate
  dim cboVal
  cboVal = Request.Form("cboFacility")
  
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************writting the SQL ******************************
      
  '------------------------get the faculty for the login ---------------
  strSQL = "Select * "_
  &" from tblfacilitySupervisor,tblFaculty "_
  &" where tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId "_
  &" and tblFacilitySupervisor.strLoginId = '"& loginId &"'" 
  
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  'Response.Write(strSQL) 
  rsSearchFaculty.Open strSQL, Conn, 3, 3     
  strFacultyName = rsSearchFaculty(7)     
  strGivenName = rsSearchFaculty(3)
  strSurname = rsSearchFaculty(4)
  strName = cstr(strGivenName) + " " + cstr(strSurname)
  %><br>
  <font size="4" face="Tahoma"> 
<img border="0" src="utslogo.gif" width="184" height="41"> <font face="Tahoma">
<b>EH</b> 
&amp; <b>S</b> <b>R</b>isk <b>R</b>egister Sorted by <b>Date </b>for <b>Faculty</b> 
, <b>Supervisor</b> , <b>Facility</b><p>&nbsp;<b>Name of Supervisor : <%Response.Write (strName)%><br>
&nbsp;Name of Faculty / Unit : <%Response.Write(strFacultyName) %></b>

<%
  
   
 '*********************Writting the report**************************** 
 
 if cboVal = 0 then
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' ORDER BY dtDateCreated,strRoomName" 
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
      rsSearchH.Open strSQL, Conn, 3, 3 
  %>
  
<B> Selection Done on :</B><BR>
<B> &nbsp;All Facility Room Name/Numbers </B>     
<% else
 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strSupervisor = '"& loginId &"' ORDER BY dtDateCreated,strRoomName"
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
    rsSearchH.Open strSQL, Conn, 3, 3 %>
   
<B> Selection Done on :</B><BR>
<B> &nbsp;Facility Room Name/Number : <%=cstr(rsSearchH(18))+"/"+ cstr(rsSearchH(19))%></B>      
 <%end if 
  
 
  'Response.Write(strSQL) 
  if not rsSearchH.EOF then 
       %> 
  </p>
  <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td width="314" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="277"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21">
		<b><font color="#800000" size="2" face="Tahoma">Risk Level</b></td>
		<td height="21" width="111">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
	<%
 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",5,rsSearchH(6))
    %>
    <tr>
		<td width="314" bgcolor="#C0C0C0"><% Response.Write(rsSearchH(8))%></td>
		<td width="277"><%=cstr(rsSearchH(18))+"/"+ cstr(rsSearchH(19))+","+ cstr(rsSearchH(23))+","+ cstr(rsSearchH(26)) %></td>
		<td bgcolor="#C0C0C0"><%=rsSearchH(9)%></td>
		<td width="111"><%=dtDate%></td>
	</tr>

<%
    rsSearchH.MoveNext  
 wend 
 %>
 </table>
 <%
 end if %>
  
<body>
</p>
</body>
</html>