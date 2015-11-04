<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">

<%
'Campbells borrowed code to escape the output 30/6/2006
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
Escape = strReturn
End Function

'*********************declaring the variables************************

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
  dim numPageStatus
  
    numPageStatus = request.querystring("cboValDummy")
    numOptionId = request.querystring("numOptionId")   
      
   ' Response.Write(numOptionId) 
    
     if numPageStatus = "1" then 
         cboVal = session("cboVal")
         cboval = cint(cboVal)
        
    else
       cboVal = Request.Form("cboFacility")  
       session("cboVal")= cboVal
      ' Response.Write (Session("cboVal")) 
     end if  
  
loginId = session("LoginId")
FacultyId = session("facultyId")

'cboVal = Request.Form("cboFacility")
'session("facilityID") = cboVal

' Response.Write(facultyID)
' Response.Write("login : "+loginID)
' Response.Write(cboVal)
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<script type="text/javascript" src="sorttable.js"></script>
<title>Online Risk Register - Report for Administrator Login</title>
</head>
<body>

<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">My Risk Assessments</h1>

<center>

<%
  
  ' Analyse the case of the input and then navigate to the proper function
  
  'case 1 : only faculty
  'case 2 : faculty , login and facility
  
     if facultyID <> 0 and loginId = "0" and cboVal = 0 then
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
  
				  strSQL = "Select * from tblFaculty where numFacultyId = "& FacultyID 
				  'Response.Write(strSQL)
				'Response.Write("test dlj")
				'Response.Write(FacultyID)
				  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
				  rsSearchFaculty.Open strSQL, Conn, 3, 3     
				  strFacultyName = rsSearchFaculty("strFacultyName")     
  
				  %>
				  
				  <table class="suprreportheader">
				<tr>
				 <th>Faculty/Unit:</th>
				 <td><%Response.Write(strFacultyName) %></td>
				</table>
				
				<br />
			<%	
			
			      strSQL = "SELECT * FROM tblQORA, tblFacility, tblBuilding, tblCampus "_
						 &" WHERE  numFacultyId = "& FacultyID & " and "_
						 &" tblQORA.numFacilityId = tblFacility.numFacilityID and "_
						 &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
						 &" tblQORA.numCampusId = tblCampus.numCampusID  "_
						 &" ORDER BY strTaskDescription"
						 
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
								<table class="sortable suprlevel" id="id12" style="width:95%;" >
								<thead>
									<tr>
										<th class="haztaskresult">Hazardous Task</th>
										<th class="assochazards">Associated Hazards</th>
										<th class="currentcontrols">Current Controls</th>
										<th class="risklevel">Risk Level</th>
										<th class="location">Location</th>
										<th class="furtheraction">Further Action required</th>
										
										<th class="dac">Date Actions Completed</th>
										<th class="renewaldate">Renewal Date</th>
										</tr>
										</thead>
										<caption>Click a table heading to sort by the respective criteria.  To edit a risk assessment, click on its title under the &quot;Hazardous Task&quot; heading.</caption>
									<%
								  while not rsSearchH.EOF 
								    dtDate = dateAdd("yyyy",2,rsSearchH(6))
								   
								    'Response.Write("Exception caught")
											%>
											<tbody>
											<tr>
												<td><a target="Operation" href="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>" title="Click to edit this QORA."><% Response.Write(rsSearchH(8))%>&nbsp;</td>
												<!-- <td><% Response.Write(rsSearchH(11))%></td> -->
												<td><%=rsSearchH(11)%></td>
												<td><% Response.Write(rsSearchH(10))%></td>
												<td><%=rsSearchH(9)%></td>
												<td><%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %></td>
												<td><% Response.Write(rsSearchH(15))%><BR><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br />
											<%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br />    
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
											Assessment</a> <%end if%></td>
												
												<td><% Response.Write(rsSearchH(17))%></td>
												<td><%=dtDate%></td>
											</tr>

								<%
								     
								    rsSearchH.MoveNext  
								 wend 
								 %>
								</tbody>
								</table>
								 <%
							end if 


'*****************************REPORT FOR ONLY FACULTY SELECTION ENDS HERE************************%>				
<%
				   
'*************************************************************************************************
case "2" : ' Selection for everything.

'*************************************************************************************************
'strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
'&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
'&" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
'&" tblQORA.numCampusId = tblCampus.numCampusID and "_
'&" tblQORA.strsupervisor = tblfacilitySupervisor.strLoginID and "_
'&" tblQORA.numFacilityId = "& cboVal &" and tblQORA.numfacultyId = "& FacultyId &" and "_
'&" strSupervisor = '"& loginId &"'  ORDER BY strAssessRisk,dtDateCreated,strRoomName"
	
	'DLJ had to edit above sql to use strSupervisor in tblFacility rather than strSupervisor in tblQORA
	'strSupervisor is redundant and is not updated by system. It should not be used.

' AA Jan 2010 relatonship fix: altered this line
'&" tblFacility.strFacilitySupervisor = tblFacilitySupervisor.strLoginID and "_
strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
&" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
&" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
&" tblQORA.numCampusId = tblCampus.numCampusID and "_
&" tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and "_
&" tblQORA.numFacilityId = "& cboVal &" and tblQORA.numfacultyId = "& FacultyId &" and "_
&" strLoginID = '"& loginId &"'  ORDER BY strAssessRisk,dtDateCreated,strRoomName"




							  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
							 ' Response.Write(strSQL) 
							  rsSearchFaculty.Open strSQL, Conn, 3, 3     
							  strFacultyName = rsSearchFaculty(19)     
							  strGivenName = rsSearchFaculty(32)
							  strSurname = rsSearchFaculty(33)
							  strName = cstr(strGivenName) + " " + cstr(strSurname)
							  %><br />
							   
							
							<table class="suprreportheader">
							<tr>
							 <th>Faculty/Unit:</th>
							 <td><%Response.Write(strFacultyName) %></td>
							</tr>
							<tr>
							 <th>Supervisor:</th>
							 <td><%Response.Write (strName)%></td>
							</tr>
						 <%if not rsSearchfaculty.EOF then%>
							<tr>
							 <th>Location:</th>
							 <td><%=cstr(rsSearchFaculty(36))+","+ cstr(rsSearchFaculty(39)) %></td>
							</tr>
							<tr>
							 <th>Facility Room Name/Number:</th>
							 <td><%=cstr(rsSearchFaculty(25))+"/"+ cstr(rsSearchFaculty(26))%></td>
							</tr>
						 <%end if%>
						<%
			select case numOptionId
			   
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
							</table>

							<br />

							<table width="100%" class="sortable suprlevel" id="id132">
							<thead>
							<tr>
							 <th class="haztaskresult">Hazardous Task</th>
							 <th class="assochazards">Associated Hazards</th>
							 <th class="currentcontrols">Current Controls</th>
							 <th class="risklevel">Risk Level</th>
							 <th class="furtheraction">Further Action Required</th>
							 
							 <th class="dac">Date Actions Completed</th>
							 <th class="renewaldate">Renewal Date</th>
							</tr>
							</thead>
							<caption>Click a table heading to sort by the respective criteria.  To edit a risk assessment, click on its title under the &quot;Hazardous Task&quot; heading.</caption>
								<%
							  while not rsSearchH.EOF 
							    dtDate = dateAdd("yyyy",2,rsSearchH(6))
							    
											'Response.Write("Expection caught")
											%>
											<tr>
												<td><a target="Operation" href="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>" title="Edit this QORA."><% Response.Write(rsSearchH(8))%></td>
												<td><%=Escape(rsSearchH(11))%></td>
												<td><%=Escape(rsSearchH(10))%></td>
												<td><%=rsSearchH(9)%></td>
												<td><%=Escape(rsSearchH(15))%><BR><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br />
											<%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br />    
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
											Assessment</a> <%end if%></td>
												
												<td><% Response.Write(rsSearchH(17))%></td>
												<td><%=dtDate%></td>
											</tr><%
							     
							      rsSearchH.MoveNext  
							   wend 
							   end if 
'***********************************************************************************************************************
	    end select
'***********************************************************************************************************************
%>
 </table>

</center>

</div></div>

</body>
</html>