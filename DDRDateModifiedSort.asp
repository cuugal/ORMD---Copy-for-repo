<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>

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
  dim loginId
  dim facultyID
  dim caseval
  dim numOptionId
  
loginId = Session("LoginId")
FacultyId = Session("facultyId")
numOptionId = Request.QueryString("numOptionID")

cboVal = session("cboVal")
cboVal = cint(cboVal)

 'Response.Write(facultyID)
 'Response.Write("login : "+loginID)
 'Response.Write(cboVal)
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Risk Level Report for Supervisors</title>
</head>
<br>
<%
  
  ' Analyse the case of the input and then navigate to the proper function
  
  'case 1 : only faculty
  'case 2 : faculty , login and facility
  
     if facultyID <> 0 and len(loginId) > 1 and cboVal = 0 then
       
          caseval = 1
     else 
       if  facultyID <> 0 and len(loginId)>1 and cboVal <> 0 then
          caseval = 2
       end if
     end if
 
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************writting the SQL ******************************
  '------------------------get the faculty for the login ---------------
          select case caseval 
  '************************************************************************************************
          case "1" :  ' Only for Faculty selection
          
  '***************************** REPORT FOR ONLY FACULTY SELECTION********************************* 	
  
				    strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
                          &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
                          &" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
                          &" tblQORA.numCampusId = tblCampus.numCampusID and "_
                          &" tblQORA.strsupervisor = tblfacilitySupervisor.strLoginID and "_
                          &" tblQORA.numfacultyId = "& FacultyId &" "_
                          &" ORDER BY strAssessRisk,dtDateCreated,strRoomName"
                          
				  'Response.Write(strSQL)
				  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
				  rsSearchFaculty.Open strSQL, Conn, 3, 3     
				  strFacultyName = rsSearchFaculty("strFacultyName")     
  
				  %><br>
				  <font size="4" face="Tahoma"> 
				<img border="0" src="utslogo.gif" width="184" height="41">Faculty/Unit: <%Response.Write(strFacultyName) %><br>
				
					<%	
		'*******************************nested case for the sort function******************************
			   select case numOptionId
			   
			   case "1":
			      		
					     strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
						 &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
						 &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
						 &" tblQORA.numCampusId = tblCampus.numCampusID  "_
						 &" ORDER BY strTaskDescription"
						 
				case "2":
			      		
					     strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
						 &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
						 &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
						 &" tblQORA.numCampusId = tblCampus.numCampusID  "_
						 &" ORDER BY strAssessRisk"
				
				case "3":
			      		
					     strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
						 &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
						 &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
						 &" tblQORA.numCampusId = tblCampus.numCampusID  "_
						 &" ORDER BY strDateActionsCompleted"
				
				case "4":
			      		
					     strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
						 &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
						 &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
						 &" tblQORA.numCampusId = tblCampus.numCampusID  "_
						 &" ORDER BY dtDateCreated"
						
						 
              end select     
              		         
							  set rsSearchH = server.CreateObject("ADODB.Recordset")
							  rsSearchH.Open strSQL, Conn, 3, 3 %>
   
							<% if not rsSearchH.EOF then 
								       %> 
								  </p>
								  <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
									<tr>
										<td><a href="DDRDateModified.asp?numOptionId=1&cboValDummy=0">Hazardous Task</a></td>
										
										<td>Hazards</td>
										<td>Controls</td>
										<td><a href="DDRDateModified.asp?numOptionId=2&cboValDummy=0">Risk Level</a></td>
										<td>Location</td>
										<td>Further Actions</td>
										<td>Comments</td>
										<td><a href="DDRDateModified.asp?numOptionId=3&cboValDummy=0">Date Actions Completed</a></td>
										<td><a href="DDRDateModified.asp?numOptionId=4&cboValDummy=0">Renewal Date</a></td>
									</tr>
									<%
								  while not rsSearchH.EOF 
								    dtDate = dateAdd("yyyy",2,rsSearchH(6))
								    if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
								     if rsSearchH(17) ="blank" then
								    'Response.Write("Expection caught")
											%>
											<tr>
												<td width="314" bgcolor="#C0C0C0"><% Response.Write(rsSearchH(8))%></td>
												<td><% Response.Write(rsSearchH(11))%></td>
												<td><% Response.Write(rsSearchH(10))%></td>
												<td><%=rsSearchH(9)%></td>
												<td><%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %></td>
												<td><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br>
											<%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br>    
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
											Assessment</a> <%end if%></td>
												<td width="111"><% Response.Write(rsSearchH(15))%></td>
												<td width="111"><% Response.Write(rsSearchH(17))%></td>
												<td width="111"><%=dtDate%></td>
											</tr>

								<%
								        end if 
								     end if
								    rsSearchH.MoveNext  
								 wend 
								 %>
								 </table>
								 <%
							end if 


'*****************************REPORT FOR ONLY FACULTY SELECTION ENDS HERE************************%>				
<%
				   
'*************************************************************************************************
case "2" : ' Selection for everything.

'*************************************************************************************************
   strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
							  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
							  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
							  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
							  &" tblQORA.strsupervisor = tblfacilitySupervisor.strLoginID and "_
							  &" tblQORA.numFacilityId = "& cboVal &" and tblQORA.numfacultyId = "& FacultyId &" and "_
							  &" strSupervisor = '"& loginId &"'  ORDER BY strAssessRisk,dtDateCreated,strRoomName"
							  
							  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
							 ' Response.Write(strSQL) 
							  rsSearchFaculty.Open strSQL, Conn, 3, 3 
							      
							  strFacultyName = rsSearchFaculty(19)     
							  strGivenName = rsSearchFaculty(32)
							  strSurname = rsSearchFaculty(33)
							  strName = cstr(strGivenName) + " " + cstr(strSurname)
							  %><br>
							  <font size="4" face="Tahoma"> 
							<img border="0" src="utslogo.gif" width="184" height="41"> <font face="Tahoma">
							EH 
							&amp; S Risk Register Sorted by Date for Faculty 
							, Supervisor , Facility<p>&nbsp;Name of Supervisor : <%Response.Write (strName)%><br>
							&nbsp;Name of Faculty / Unit : <%Response.Write(strFacultyName) %><br>
                              
                       		 <%if not rsSearchfaculty.EOF then%> &nbsp;Facility Room Name/Number : <%=cstr(rsSearchFaculty(25))+"/"+ cstr(rsSearchFaculty(26))%>      
						 <br>&nbsp;Location :  <%=cstr(rsSearchFaculty(36))+","+ cstr(rsSearchFaculty(39)) %>
						 <%end if%>
						<%
						
			select case numOptionID		
		       case "1":
		          
			      		
					 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
							  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
							  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
							  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
							  &" tblQORA.numFacilityId = "& cboVal &" and "_
							  &" strSupervisor = '"& loginId &"'  ORDER BY strTaskDescription"
						 
				case "2":
			      		
					  strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
							  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
							  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
							  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
							  &" tblQORA.numFacilityId = "& cboVal &" and "_
							  &" strSupervisor = '"& loginId &"'  ORDER BY strAssessRisk"
				case "3":
			      		strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
							  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
							  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
							  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
							  &" tblQORA.numFacilityId = "& cboVal &" and "_
							  &" strSupervisor = '"& loginId &"'  ORDER BY strDateActionsCompleted"
				
				case "4":
			      		
					 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
							  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
							  &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
							  &" tblQORA.numCampusId = tblCampus.numCampusID and "_
							  &" tblQORA.numFacilityId = "& cboVal &" and "_
							  &" strSupervisor = '"& loginId &"'  ORDER BY dtDateCreated"
              end select 
              
						  
						  'Response.Write(strSQL)
						    set rsSearchH = server.CreateObject("ADODB.Recordset")
						    rsSearchH.Open strSQL, Conn, 3, 3 %>
						   
							 <%

							  'Response.Write(strSQL) 
						if not rsSearchH.EOF then 
							       %> 
							  </p>
							  <table border="2" width="85%" id="table1" bordercolor="#FFFFFF">
								<tr>
									<td>
									
									<a href="DDRDateModified.asp?numOptionId=1&cboValDummy=1">
									Hazardous Task</a></td>
									<td>
										
										Hazards</td>

									<td>
									Controls</td>
								
									<td>
									
									<a href="DDRDateModified.asp?numOptionId=2&cboValDummy=1">
									Risk Level</a></td>
									<td>
									Further Actions</td>
									<td>
									Comments</td>
									<td>
									
									<a href="DDRDateModified.asp?numOptionId=3&cboValDummy=1">
									Date Actions Completed</a></td>
									<td>
									
									<a href="DDRDateModified.asp?numOptionId=4&cboValDummy=1">
									Renewal Date</a></td>
								</tr>
								<%
							  while not rsSearchH.EOF 
							    dtDate = dateAdd("yyyy",2,rsSearchH(6))
							    if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
							     if rsSearchH(17) ="blank" then
											'Response.Write("Expection caught")
											%>
											<tr>
												<td width="314" bgcolor="#C0C0C0"><% Response.Write(rsSearchH(8))%></td>
												<td><% Response.Write(rsSearchH(11))%></td>
												<td><% Response.Write(rsSearchH(10))%></td>
												
												<td><%=rsSearchH(9)%></td>
												<td>  <%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br>
											<%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br>    
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
											Assessment</a> <%end if%></td>
												<td width="111"><% Response.Write(rsSearchH(15))%></td>
												<td width="111"><% Response.Write(rsSearchH(17))%></td>
												<td width="111"><%=dtDate%></td>
											</tr><%
							       end if 
							     end if
							      rsSearchH.MoveNext  
							   wend 
							   end if 
'***********************************************************************************************************************							   
	    end select
'***********************************************************************************************************************							 %>
							 </table>
<body link="#800000" vlink="#800000" alink="#800000">
</p>
</body>
</html>