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
  'cboVal = Request.Form("cboFacility")
  cboVal = Request.QueryString("cboVal")
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
  '********************************************************************
  dim numOptionId
  	 
if cboVal = 0 then
 numOptionId = Request.QueryString("numOptionID")
 'Response.Write(numOptionId) 
			
 
 '*********************Writting the report****************************   %>
</p>
<p>&nbsp;<B>Selection Done on :&nbsp;<br>
&nbsp;All Facility Room Name/Numbers </B> 
       <%
       strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strTaskDescription" %>
				<BR><b>&nbsp;Click on one of the headings to sort by that criteria.</b><table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
			<td width="314" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">
			Hazardous Task</b></td>
				<td width="314" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">
			Hazards</b></td>
		<td height="21" width="277"><b>
		<font size="2" face="Tahoma" color="#800000">
		Controls</b></td>
		<td height="21" width="277"><b><font face="Tahoma" size="2" color="#800000">
		<a href="SupRDateModified.asp?numOptionId=3&cboValDummy=0">Location</a></b></td>
		<td height="21">
		<b><font color="#800000" size="2" face="Tahoma">
		<a href="SupRDateModified.asp?numOptionId=4&cboValDummy=0">Risk Leve</a>l</b></td>
		<td height="21" width="66">
		<b><font size="2" face="Tahoma" color="#800000">Further Actions</b></td>
		<td height="21" width="67">
		<b><font size="2" face="Tahoma" color="#800000">Comments</b></td>
		<td height="21" width="111">
		<b><font color="#800000" size="2" face="Tahoma">
		<a href="SupRDateModified.asp?numOptionId=5&cboValDummy=0">Date Actions Completed</a></b></td>
		<td height="21" width="111">
		<font color="#800000" size="2" face="Tahoma"><b>
		<a href="SupRDateModified.asp?numOptionId=6&cboValDummy=0">Renewal Date</a></b></td>
	</tr>
	<%
	
	 

select case numOptionId
case "1" :

				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strTaskDescription" 
case "2" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strControlRiskDesc" 
				
case "3" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strRoomName,strRoomNumber,strBuildingName,strCampusName" 

case "4" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strAssessRisk" 

case "5" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strDateActionsCompleted" 

case "6" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY dtDateCreated" 

end select
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
      rsSearchH.Open strSQL, Conn, 3, 3 
      
      'Response.Write(strSQL) 
  if not rsSearchH.EOF then 
       
	  
 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",5,rsSearchH(6))
    %>
    <tr>
		<td width="314" bgcolor="#C0C0C0"><A target ="Operation" HREF ="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>"><% Response.Write(rsSearchH(8))%></td>
		<td width="277"><% Response.Write(rsSearchH(11))%></td>
		<td width="277"><% Response.Write(rsSearchH(10))%></td>
		<td width="277"><%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %></td>
		<td bgcolor="#C0C0C0"><%=rsSearchH(9)%></td>
		<td width="66"><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><BR><BR>
    <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a> <%end if%><BR><BR>    
    <%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
    Assessment</a> <%end if%></td>
		<td width="67"><% Response.Write(rsSearchH(15))%></td>
		<td width="111"><% Response.Write(rsSearchH(17))%></td>
		<td width="111"><%=dtDate%></td>
	</tr>

<%
    rsSearchH.MoveNext  
 wend 
 %>
 </table>
 <%
 end if %>

    
<% else

 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strSupervisor = '"& loginId &"' ORDER BY dtDateCreated,strRoomName"
  
  numOptionId = Request.QueryString("numOptionID")
 'Response.Write(numOptionId) 
			
select case numOptionId
case "1" :

				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strTaskDescription" 
case "2" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strControlRiskDesc" 
				
case "3" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strRoomName,strRoomNumber,strBuildingName,strCampusName" 

case "4" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strAssessRisk" 

case "5" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY strDateActionsCompleted" 

case "6" :
				strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
				&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
				&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
				&" tblQORA.numCampusId = tblCampus.numCampusID and "_
				&" strSupervisor = '"& loginId &"' ORDER BY dtDateCreated" 

end select
  
    set rsSearchH = server.CreateObject("ADODB.Recordset")
    rsSearchH.Open strSQL, Conn, 3, 3 %>
   
<BR><BR><B> &nbsp;Selection Done on :</B><BR>
<B> &nbsp;Facility Room Name/Number : <%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))%></B>      
 </p>&nbsp;<b>Location :  <%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %></b>
 <%end if  
  'Response.Write(strSQL) 
  if not rsSearchH.EOF then 
       %> 
  <BR><b>&nbsp;Click on one of the headings to sort by that criteria.</b><table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
		<tr>
		<td width="170" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">
		Hazardous Task</b></td>
		<td width="170" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">
		Hazards</b></td>
		<td height="21" width="152"><b>
		<font size="2" face="Tahoma" color="#800000">
		Controls</b></td>
	
		<td height="21">
		<b><font color="#800000" size="2" face="Tahoma">
		<a href="SupRDateModified.asp?numOptionID=4&cboValDummy=1">Risk Level</a></b></td>
		<td height="21" width="50">
		<b><font size="2" face="Tahoma" color="#800000">Further Actions</b></td>
		<td height="21" width="67">
		<b><font size="2" face="Tahoma" color="#800000">Comments</b></td>
		<td height="21" width="111">
		<b><font color="#800000" size="2" face="Tahoma">
		<a href="SupRDateModified.asp?numOptionID=5&cboValDummy=1">Date Actions Completed</a></b></td>
		<td height="21" width="111">
		<font color="#800000" size="2" face="Tahoma"><b>
		<a href="SupRDateModified.asp?numOptionID=6&cboValDummy=1">Renewal Date</a></b></td>
	</tr>
	<%
 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",5,rsSearchH(6))
    %>
    <tr>
		<td width="170" bgcolor="#C0C0C0"><A target ="Operation" HREF ="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>"><% Response.Write(rsSearchH(8))%></td>
        <td width="152"><% Response.Write(rsSearchH(11))%></td>
		<td width="152"><% Response.Write(rsSearchH(10))%></td>
		
		<td bgcolor="#C0C0C0"><%=rsSearchH(9)%></td>
		<td width="50"><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><BR>
    <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a> <%end if%><BR>    
    <%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
    Assessment</a> <%end if%></td>
		<td width="67"><% Response.Write(rsSearchH(15))%></td>
		<td width="111"><% Response.Write(rsSearchH(17))%></td>
		<td width="111"><%=dtDate%></td>
	</tr>

<%
    rsSearchH.MoveNext  
 wend 
 %>
 </table>
 <%
 end if %>
  
<body link="#800000" vlink="#800000" alink="#800000">
</p>
</body>
</html>