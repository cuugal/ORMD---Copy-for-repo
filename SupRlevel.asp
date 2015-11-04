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
  dim flg1
  dim flg2
  dim flg3  
  
   flg1 = 0
   flg2 = 0
   flg3 = 0
   
  cboVal = Request.Form("cboFacility")
  
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************writting the SQL ******************************
  
  '------------------------get the faculty for the login ---------------
  strSQL = "Select * "_
  &" from tblfacilitySupervisor,tblFaculty"_
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
<img border="0" src="utslogo.gif" width="184" height="41"> <b>&nbsp;</b><font face="Tahoma"><b>EH</b> 
&amp; <b>S</b> <b>R</b>isk <b>R</b>egister Sorted by <b>Risk </b>for <b>Faculty</b> 
, <b>Supervisor</b> , <b>Facility</b><p><b>&nbsp;</b><b>Name of Supervisor : <%Response.Write (strName)%><br>
&nbsp;Name of Faculty / Unit : <%Response.Write(strFacultyName) %></b>
<BR>
<%
  
   
 '*********************Writting the report**************************** 
 ' code for sorting the risk levels = H
 
 
  if cboVal = 0 then    
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and strAssessRisk = 'H' ORDER BY strRoomName"
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
      rsSearchH.Open strSQL, Conn, 3, 3 %>
       
<B> Selection Done on :</B><BR>
<B> &nbsp;All Facility Room Name/Numbers </B>     
      
<%  else
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and "_
  &" tblQORA.numfacilityId = "& cboVal &" and"_
  &" strAssessRisk = 'H' ORDER BY strRoomName"
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
      rsSearchH.Open strSQL, Conn, 3, 3
      
       set rsFillF = server.CreateObject("ADODB.Recordset")
      rsFillF.Open "Select * from tblFacility where numFacilityID = "& cboVal &"", Conn, 3, 3    %>
   
<B> Selection Done on :</B><BR>
<B> &nbsp;Facility Room Name/Number : <%=cstr(rsFillF(1))+"/"+ cstr(rsFillF(2))%></B>   
 <% end if %>

 
  <%
  if not rsSearchH.EOF then 
  flg1 =1
   %></p>
  <table border="2" width="75%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td width="370" bgcolor="#FFFF99"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High</b></td>
		<td width="240">&nbsp;</td>
		<td bgcolor="#FFFF99"><b><font face="Tahoma" size="2" color="#800000">&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="370" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="240"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
	<%
 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",5,rsSearchH(6)) %>
    <tr>
		<td width="370" bgcolor="#C0C0C0"><% Response.Write(rsSearchH(8))%></td>
		<td width="240"><%=cstr(rsSearchH(18))+"/"+ cstr(rsSearchH(19))+","+ cstr(rsSearchH(23))+","+ cstr(rsSearchH(26)) %></td>
		<td bgcolor="#C0C0C0"><%=dtDate%></td>
	</tr>

<%
    rsSearchH.MoveNext  
 wend 
 %>
 </table>
 <%
 else
 flg1 = 0
 
 %>
   
 <%
 end if
 
 '-----------------------------------------------------------------------
 'code for sorting the risk levels = M
 
  if cboVal = 0 then
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and strAssessRisk = 'M' ORDER BY strRoomName"
  else
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and "_
  &" tblQORA.numfacilityId = "& cboVal &" and"_
  &" strAssessRisk = 'M' ORDER BY strRoomName"
  end if
  
  set rsSearchM = server.CreateObject("ADODB.Recordset")
      rsSearchM.Open strSQL, Conn, 3, 3 
  
      
      if not rsSearchM.EOF then
      flg2 = 1
  %><BR>
 <table border="2" width="75%" id="table1" bordercolor="#FFFFFF">
 	<tr>
		<td width="370" bgcolor="#FFFF99"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium</b></td>
		<td width="243">&nbsp;</td>
		<td bgcolor="#FFFF99"><b><font face="Tahoma" size="2" color="#800000">&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="370" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="243"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
 <%  
 while not rsSearchM.EOF
  dtDate = dateAdd("yyyy",5,rsSearchM(6)) %>
  <tr>
		<td width="370" bgcolor="#C0C0C0"><% Response.Write(rsSearchM(8))%></td>
		<td width="243"><%=cstr(rsSearchM(18))+"-"+ cstr(rsSearchM(19))+","+ cstr(rsSearchM(23))+","+ cstr(rsSearchM(26)) %></td>
		<td bgcolor="#C0C0C0"><%=dtDate%></td>
	</tr>
   <%
    rsSearchM.MoveNext  
 wend 
 %>
 </table>
  <%
 else
 flg2 = 0
 
 %>
 <%end if
 
 '-----------------------------------------------------------------------
 'code for sorting the risk levels = L
 
 if cboVal = 0 then
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and strAssessRisk = 'L' ORDER BY strRoomName"
  else
  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" strSupervisor = '"& loginId &"' and "_
  &" tblQORA.numfacilityId = "& cboVal &" and"_
  &" strAssessRisk = 'L' ORDER BY strRoomName"
  end if
  set rsSearchL = server.CreateObject("ADODB.Recordset")
      rsSearchL.Open strSQL, Conn, 3, 3 
      
      if not rsSearchL.EOF then
      flg3 =1
  %><BR>
 <table border="2" width="75%" id="table1" bordercolor="#FFFFFF">
 	<tr>
		<td width="370" bgcolor="#FFFF99"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L - Low</b></td>
		<td width="241">&nbsp;</td>
		<td bgcolor="#FFFF99"><b><font face="Tahoma" size="2" color="#800000">&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="370" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="241"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
 <%  
 while not rsSearchL.EOF 
 dtDate = dateAdd("yyyy",5,rsSearchL(6))%>
  <tr>
		<td width="370" bgcolor="#C0C0C0"><% Response.Write(rsSearchL(8))%></td>
		<td width="241"><%=cstr(rsSearchL(18))+"-"+ cstr(rsSearchL(19))+","+ cstr(rsSearchL(23))+","+ cstr(rsSearchL(26)) %></td>
		<td bgcolor="#C0C0C0"><%=dtDate%></td>
	</tr>
   <%
    rsSearchL.MoveNext  
 wend 
 %>
 </table>
  <%
 else
 flg3 = 0
 
 %>
 <%end if%>
 <%
  if flg1 = 0 and flg2 = 0 and flg3 = 0 then  %>
   <BR>
    
   <B> Record not present under this section !</B>
  <%end if%>
<body>
</p>
</body>
</html>