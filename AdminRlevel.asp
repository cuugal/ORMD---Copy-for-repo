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
<p><img border="0" src="utslogo.gif" width="184" height="41">&nbsp; <b>&nbsp;</b><font size="4" face="Tahoma"><font face="Tahoma"><b>EH</b> 
&amp; <b>S</b> <b>R</b>isk <b>R</b>egister Sorted by <b>Risk </b>
<b><font face="Tahoma" size="4">&nbsp;</b></p>
<%'*********************declaring the variables************************
   
  dim rsSearchH
  dim rsSearchM
  dim rsSearchL 
  dim rsFillFacultyH
  dim rsFillLocation
  dim rsSearchFaculty
  dim Conn 
  dim strSQL
  dim strFacultyName
  dim strGivenName
  dim strSurname
  dim strName
  dim dtDate
  dim FacultyVal
  dim SupervisorVal
  dim FacilityVal
  dim caseVal
  dim A
  dim B
  dim C
  
  FacultyVal = Request.Form("hdnfaculty")
  FacilityVal = Request.Form("cboFacility")
  SupervisorVal = cstr(Request.Form("hdnSupervisor")) 
  
  'SupervisorVal = len(SupervisorVal)
  'Response.Write(SupervisorVal)
  
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************creating the conditions for SQL *******************************
if FacultyVal <> 0 then
	if len(SupervisorVal) > 1 then
		caseVal = 1 ' sql for a particular supervisor
	else
	    caseVal = 2 ' sql for a particular faculty (= sql for all supervisors)
	end if
else
       caseVal = 3 ' sql for All faculties        
end if

if len(SupervisorVal) > 1   then
  caseval = 4 ' sql for a particular supervisor
'else
'  caseVal = 5 ' sql for all supervisor    
end if 


if FacilityVal <> 0 then 
   caseval = 5 ' sql for a particular facility 
'else
'   caseVal = 7 ' sql for all facilitiespervisor
end if	

if FacilityVal <> 0 and len(SupervisorVal) > 1 then
   caseVal = 6 ' sql for a particular facility and su
end if
	
'**************************** select case for the SQLs

select case caseVal
	
	case "4":'sql for a particular supervisor
%><B>	Selection Done on : </B>
<BR>	
<%	      if FacultyVal <> 0 then
	      set rsFacname = server.CreateObject("ADODB.Recordset")
			 rsFacname.Open "Select * from tblfaculty where numFacultyId = "& FacultyVal &"", Conn, 3, 3%>
	   
 <B>Faculty Name / Unit :  <%=(rsFacname(1))%></B><BR>
	     <% ' code for the addition of faculty name from the selection criterion
	      end if 
	      
	  			 set rsFullName2 = server.CreateObject("ADODB.Recordset")
			 rsFullName2.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& SupervisorVal &"'", Conn, 3, 3%>
	   
 <B>Supervisor Name :  <%=cstr(rsFullName2(0)) + " " + cstr(rsFullName2(1))%></B>
			
			<%strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.strSupervisor = '"& SupervisorVal &"' and"_
			&" strAssessRisk = 'H' ORDER BY strRoomName"
	         
			 set rsSearchFacultyH = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyH.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyH.EOF then
			   A = true  
			   
			      'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
%>
 <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyH.EOF  
   'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyH(6)) %> 
	<tr>	
		<td><% Response.Write(rsSearchFacultyH(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyH(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyH(18))+"/"+ cstr(rsSearchFacultyH(19))+","+ cstr(rsSearchFacultyH(23))+","+ cstr(rsSearchFacultyH(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyH.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   medium risk code----------------------------------------------
  
  strSQL ="SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.strSupervisor = '"& SupervisorVal &"' and"_
			&" strAssessRisk = 'M' ORDER BY strRoomName"
	         
			 set rsSearchFacultyM = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyM.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyM.EOF then 
			  B = true %>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium &nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyM.EOF 
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyM(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3 
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyM(6)) %> 
	<tr>	
		<td><% Response.Write(rsSearchFacultyM(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyM(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyM(18))+"/"+ cstr(rsSearchFacultyM(19))+","+ cstr(rsSearchFacultyM(23))+","+ cstr(rsSearchFacultyM(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyM.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   Low risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.strSupervisor = '"& SupervisorVal &"' and"_
			&" strAssessRisk = 'L' ORDER BY strRoomName"
	         
			 set rsSearchFacultyL = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyL.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyL.EOF then 
			 c = true %>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L- Low&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyL.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyL(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyL(6)) %> 
	<tr>	
		<td><% Response.Write(rsSearchFacultyL(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyL(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyL(18))+"/"+ cstr(rsSearchFacultyL(19))+","+ cstr(rsSearchFacultyL(23))+","+ cstr(rsSearchFacultyL(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyL.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
 <BR><% if A = False and B = false and C = False then
           Response.Write("Records Not Present !")
        end if
 
'***************************************************************************************************************				
	case "2":'sql for a particular faculty (= sql for all supervisors) 
	%>
	          <B>	Selection Done on : </B>
<BR>	
	  <%    set rsFacname = server.CreateObject("ADODB.Recordset")
			 rsFacname.Open "Select * from tblfaculty where numFacultyId = "& FacultyVal &"", Conn, 3, 3%>
	   
 <B>Faculty Name / Unit :  <%=(rsFacname(1))%></B><BR>	
   
 
<%
	        strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacultyId = "& FacultyVal &" and"_
			&" strAssessRisk = 'H' ORDER BY strRoomName"
	         
			 set rsSearchFacultyH = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyH.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyH.EOF then 
			 A = true%>
			 
			
 <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyH.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyH(6)) %> 
	<tr>	
		<td><% Response.Write(rsSearchFacultyH(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyH(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyH(18))+"/"+ cstr(rsSearchFacultyH(19))+","+ cstr(rsSearchFacultyH(23))+","+ cstr(rsSearchFacultyH(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyH.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   medium risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty  "_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacultyId = "& FacultyVal &" and"_
			&" strAssessRisk = 'M' ORDER BY strRoomName"
	         
			 set rsSearchFacultyM = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyM.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyM.EOF then 
			 B = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyM.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyM(21))
			  'Response.Write(strloginFullName) 
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyM(6)) %> 
	<tr>	
		<td><%= Response.Write(rsSearchFacultyM(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyM(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyM(18))+"/"+ cstr(rsSearchFacultyM(19))+","+ cstr(rsSearchFacultyM(23))+","+ cstr(rsSearchFacultyM(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyM.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   Low risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty  "_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacultyId = "& FacultyVal &" and"_
			&" strAssessRisk = 'L' ORDER BY strRoomName"
	         
			 set rsSearchFacultyL = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyL.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyL.EOF then 
			 C = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L - Low&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyL.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyL(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyL(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyL(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyL(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyL(18))+"/"+ cstr(rsSearchFacultyL(19))+","+ cstr(rsSearchFacultyL(23))+","+ cstr(rsSearchFacultyL(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyL.MoveNext  
    wend 
  %>
 </table>
  <%end if%><BR>
<%  if A = False and B = false and C = False then
           Response.Write("Records Not Present !")
        end if
'***************************************************************************************************************				    
	case "3":'sql for All faculties
	%>			 <B>	Selection Done on : </B>
<BR>		   
 <B>All Faculties </B><%
 
		strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" strAssessRisk = 'H' ORDER BY strRoomName"
	         
			 set rsSearchFacultyH = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyH.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyH.EOF then 
			 A = true %>

 <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyH.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyH(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyH(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyH(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyH(18))+"/"+ cstr(rsSearchFacultyH(19))+","+ cstr(rsSearchFacultyH(23))+","+ cstr(rsSearchFacultyH(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyH.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   medium risk code----------------------------------------------
  
  strSQL ="SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" strAssessRisk = 'M' ORDER BY strRoomName"
	         
			 set rsSearchFacultyM = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyM.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyM.EOF then 
			 B = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium &nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyM.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyM(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyM(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyM(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyM(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyM(18))+"/"+ cstr(rsSearchFacultyM(19))+","+ cstr(rsSearchFacultyM(23))+","+ cstr(rsSearchFacultyM(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyM.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   Low risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" strAssessRisk = 'L' ORDER BY strRoomName"
	         
			 set rsSearchFacultyL = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyL.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyL.EOF then 
			 C = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L- Low&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyL.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyL(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyL(6))%>
	<tr>	
		<td><%=Response.Write(rsSearchFacultyL(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyL(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyL(18))+"/"+ cstr(rsSearchFacultyL(19))+","+ cstr(rsSearchFacultyL(23))+","+ cstr(rsSearchFacultyL(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%   rsSearchFacultyL.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
 <BR><%if A = False and B = false and C = False then
           Response.Write("Records Not Present !")
        end if
	
'***************************************************************************************************************
	case "5":'sql for  a particular facility 
	%> 
	   			 <B>	Selection Done on : </B>
	   <%
	         set rsFaciname = server.CreateObject("ADODB.Recordset")
	 rsFaciname.Open "Select * from tblfacility where numFacilityId = "& FacilityVal &"", Conn, 3, 3%>
	   
			 
<BR>		   
 <B>Facility Room Name/Number :  <%=cstr(rsFaciname(1))+"/"+ cstr(rsFaciname(2)) %></B>
	<%
	strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_ 
			&" strAssessRisk = 'H' ORDER BY strRoomName"
	         
			 set rsSearchFacultyH = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyH.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyH.EOF then 
			 A = True %>

 <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyH.EOF 
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyH(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyH(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyH(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyH(18))+"/"+ cstr(rsSearchFacultyH(19))+","+ cstr(rsSearchFacultyH(23))+","+ cstr(rsSearchFacultyH(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyH.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   medium risk code----------------------------------------------
  
  strSQL ="SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_ 
			&" strAssessRisk = 'M' ORDER BY strRoomName"
	         
			 set rsSearchFacultyM = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyM.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyM.EOF then 
			 B = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium &nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyM.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyM(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyM(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyM(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyM(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyM(18))+"/"+ cstr(rsSearchFacultyM(19))+","+ cstr(rsSearchFacultyM(23))+","+ cstr(rsSearchFacultyM(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyM.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   Low risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_ 
			&" strAssessRisk = 'L' ORDER BY strRoomName"
	         
			 set rsSearchFacultyL = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyL.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyL.EOF then 
			  C = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L- Low&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyL.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyL(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyL(6))%>
	<tr>	
		<td><%=Response.Write(rsSearchFacultyL(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyL(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyL(18))+"/"+ cstr(rsSearchFacultyL(19))+","+ cstr(rsSearchFacultyL(23))+","+ cstr(rsSearchFacultyL(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%   rsSearchFacultyL.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
 <BR><%
     if A = False and B = false and C = False then
           Response.Write("Records Not Present !")
        end if
'***************************************************************************************************************	
    case "6":' sql for a particular facility and supervisor
    %>
    <B>Selection Done on : </B>
<BR> 
         	 <% if FacultyVal <> 0 then
	             set rsFacu = server.CreateObject("ADODB.Recordset") 
			 rsFacu.Open "Select * from tblfaculty where numFacultyId = "& FacultyVal &"", Conn, 3, 3
			 %>
			 <B>Supervisor :  <%=cstr(rsFacu(1))%></B>
	         <% end if 
 
			  if FacilityVal <> 0 then
	             set rsFac = server.CreateObject("ADODB.Recordset") 
			 rsFac.Open "Select * from tblfacility where numFacilityId = "& FacilityVal &"", Conn, 3, 3
			 %>
	         <% end if 
	        	   		
			set rsFullName1 = server.CreateObject("ADODB.Recordset")
			rsFullName1.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& SupervisorVal &"'", Conn, 3, 3
            %>
            <BR>
            
 <B>Supervisor :  <%=cstr(rsFullName1(0)) + " " + cstr(rsFullName1(1))%></B> <BR>
 <B>Facility Room Name/Number :  <%=cstr(rsfac(1))+"/"+ cstr(rsfac(2)) %></B><BR>
 <% 
			  
    
    strSQL ="SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_
			&" tblQORA.strSupervisor  = '"& SupervisorVal &"' and "_ 
			&" strAssessRisk = 'H' ORDER BY strRoomName"
	         

			 set rsSearchFacultyH = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyH.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyH.EOF then 
			  A = true
			 %> 
			  
		   
 
 <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;H - High&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyH.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyH(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyH(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyH(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyH(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyH(18))+"/"+ cstr(rsSearchFacultyH(19))+","+ cstr(rsSearchFacultyH(23))+","+ cstr(rsSearchFacultyH(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyH.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   medium risk code----------------------------------------------
  
  strSQL ="SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_
			&" tblQORA.strSupervisor  = '"& SupervisorVal &"' and "_ 
			&" strAssessRisk = 'M' ORDER BY strRoomName"
	         
			 set rsSearchFacultyM = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyM.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyM.EOF then 
			 B = true%>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;M - Medium &nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyM.EOF 
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyM(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyM(6)) %> 
	<tr>	
		<td><%=Response.Write(rsSearchFacultyM(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyM(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyM(18))+"-"+ cstr(rsSearchFacultyM(19))+","+ cstr(rsSearchFacultyM(23))+","+ cstr(rsSearchFacultyM(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%    rsSearchFacultyM.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
  
  <BR><%'-------------------   Low risk code----------------------------------------------
  
  strSQL = "SELECT * FROM tblQORA,tblFacility,tblBuilding,tblCampus,tblFaculty"_
			&" WHERE tblQORA.numFacultyId = tblFaculty.numFacultyID and "_
			&" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
			&" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
			&" tblQORA.numCampusId = tblCampus.numCampusID and "_
			&" tblQORA.numfacilityId  = "& FacilityVal &" and "_
			&" tblQORA.strSupervisor  = '"& SupervisorVal &"' and "_ 
			&" strAssessRisk = 'L' ORDER BY strRoomName"
	         
			 set rsSearchFacultyL = server.CreateObject("ADODB.Recordset")
			 rsSearchFacultyL.Open strSQL, Conn, 3, 3
			 if not rsSearchFacultyL.EOF then
			   C = true %>
			 
<table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
	<tr>
		<td colspan="5"><b>
		<font face="Tahoma" size="2" color="#800000">Risk Level&nbsp; : &nbsp;L- Low&nbsp;
		</b></td>
	</tr>
	<tr>
		<td width="28%" height="21"><b>
		<font face="Tahoma" size="2" color="#800000">Hazardous Task</b></td>
		<td height="21" width="20%">
		<font color="#800000" size="2" face="Tahoma"><b>faculty / Unit </b></td>
		<td height="21" width="28%"><b><font face="Tahoma" size="2" color="#800000">
		Location</b></td>
		<td height="21" width="9%">
		<font color="#800000" size="2" face="Tahoma"><b>Supervisor</b></td>
		<td height="21" width="11%">
		<font color="#800000" size="2" face="Tahoma"><b>Renewal Date</b></td>
	</tr>
        
 <% while not rsSearchFacultyL.EOF  
 'write an extra SQL to fetch the supervisors full Name 
			  strLoginFullName = cstr(rsSearchFacultyL(21))
			  
			 set rsFullName = server.CreateObject("ADODB.Recordset")
			 rsFullName.Open "Select strGivenName,strSurname from tblfacilitysupervisor where strLoginId = '"& strLoginFullName &"'", Conn, 3, 3
    'Response.Write(rsSearchFacultyH(6))
    dtDate = dateAdd("yyyy",5,rsSearchFacultyL(6))%>
	<tr>	
		<td><%=Response.Write(rsSearchFacultyL(8))%></td>
		<td width="20%"><%=Response.Write(rsSearchFacultyL(28))%></td>
		<td width="28%"><%=cstr(rsSearchFacultyL(18))+"-"+ cstr(rsSearchFacultyL(19))+","+ cstr(rsSearchFacultyL(23))+","+ cstr(rsSearchFacultyL(26)) %></td>
		<td width="9%"><%=cstr(rsFullName(0)) + " " + cstr(rsFullName(1))%></td>
		<td width="11%"><%=dtDate%></td>
	</tr>
<%   rsSearchFacultyL.MoveNext  
    wend 
  %>
 </table>
  <%end if%>
 <BR><%
 if A = False and B = false and C = False then
           Response.Write("Records Not Present !")
        end if
'***************************************************************************************************************    
	
end select
   'Response.Write(" caseVal  : ")
   'Response.Write(caseVal)
%>
<body>
</p>
</body>
</html>