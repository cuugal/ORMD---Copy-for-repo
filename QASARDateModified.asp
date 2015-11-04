<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
'Campbells borrowed code to escape the output 26/6/2006
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
 <title>Online Risk Register - QORA Report for Administrator</title>
 <script type="text/javascript" src="sorttable.js"></script>
</head>
<body>

<div id="wrapper">

<div id="content">


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
  
  '*** Writing the SQL ***
  '*** Get the Faculty for the login ***
          select case caseval 
  '*************************************
          case "1" :  ' Only for Faculty selection
          
  '************** REPORT FOR ONLY FACULTY SELECTION **** 	
  
				  strSQL = "Select * from tblFaculty where numFacultyId = "& FacultyID 
				  'Response.Write(strSQL)
				  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
				  rsSearchFaculty.Open strSQL, Conn, 3, 3     
				  strFacultyName = rsSearchFaculty("strFacultyName")     
  
				  %><br />

<table class="suprreportheader">
 <tr>
  <th>Faculty/Unit:</th>
  <td><%Response.Write(strFacultyName) %></td>
</tr>
</table>
<br />
<%	

	strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus "_
		   &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
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
				   
                <!--   <table border="2" width="85%" id="table1" bordercolor="#FFFFFF"> -->

								<table class="suprlevel" style="width: 95%;">
									<tr>
										<th>Hazardous Task</th>
										<th>Hazards</th>
										<th>Controls</th>
										<th>Risk Level</th>
										<th>Location</th>
										<th>Further Actions</th>
										<th>Comments</th>
										<th>Date Actions Completed</th>
										<th>Renewal Date</th>
									</tr>
									<%
								  while not rsSearchH.EOF 
								    dtDate = dateAdd("yyyy",2,rsSearchH(6))
								    if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
								    Val = len(rsSearchH(17))
								     if val <=1   then
								    'Response.Write("Expection caught")
									'end if %>
									<tr>
										<td><a target="Operation" href="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>" title="Edit this Risk Assessment."><% Response.Write(rsSearchH(8))%></a></td>
										<!-- <td><% Response.Write(rsSearchH(11))%></td> -->
										<td><%=Escape(rsSearchH(11)) %></td>
										<td><%=Escape(rsSearchH(10)) %></td>
										<td class="centrecontent"><%=rsSearchH(9)%></td>
										<td><%=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %></td>
										<td><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/internal/swms.doc" title="Safe Work Method Statement (in Microsoft Word format, 47 Kb).">Safe Work Method Statement</a> <%end if%><br />

                      <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/" title="Chemical Risk Assessment">Chemical Risk Assessment</a><%end if%><br />    
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/internal/generalriskmgt.doc" title="General risk assessment form (in Microsoft Word format, 67 Kb).">General Risk Assessment</a> <%end if%></td>
										<td><% Response.Write(rsSearchH(15))%></td>
										<td><% Response.Write(rsSearchH(17))%></td>
										<td><%=dtDate%></td>
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
							  %>
                
                <br/>

				<table class="suprreportheader">
				<tr>
				  <th>Name of Faculty/Unit:</th>
				  <td><%Response.Write(strFacultyName) %></td>
				</tr>
				<tr>
				  <th>Name of Supervisor:</th>
				  <td><%Response.Write (strName)%></td>
				</tr>

						 <%if not rsSearchfaculty.EOF then%>

				<tr>
					<th>Facility Room Name/Number:</th>
					<td><%=cstr(rsSearchFaculty(25))+"/"+ cstr(rsSearchFaculty(26))%></td>
				</tr>
				<tr>
					<th>Location:</th>
					<td><%=cstr(rsSearchFaculty(36))+","+ cstr(rsSearchFaculty(39)) %></td>
				</tr>

				<%end if%>
				</table><br />
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

							<table class="sortable suprlevel" style="width: 95%;">
								<tr>
									<th>Hazardous Task</th>
									<th>Hazards</th>
									<th>Controls</th>
									<th>Risk Level</th>
									<th>Further Actions</th>
									<th>Comments</th>
									<th>Date Actions Completed</th>
									<th>Renewal Date</th>
								</tr>
							<%
							  while not rsSearchH.EOF 
							    dtDate = dateAdd("yyyy",2,rsSearchH(6))
							    if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
							     Val = len(rsSearchH(17))
								     if val <=1   then
											'Response.Write("Expection caught")
											%>
											<tr>
												<td><a target="Operation" href="EditQORA.asp?numCQORAId=<%=rsSearchH(0)%>" title="Click to edit this Risk Assessment"><% Response.Write(rsSearchH(8))%></td>
												<td><%=Escape(rsSearchH(11))%></td>
												<td><%=Escape(rsSearchH(10)) %></td>
												<td class="centrecontent"><%=rsSearchH(9)%></td>
												<td>
												<%if rsSearchH(12)= true then %>
													<a target="_blank" href="http://www.ehs.uts.edu.au/internal/swms.doc" title="Safe Work Method Statement (in Microsoft Word format, 47 Kb).">Safe Work Method Statement</a> 
						                         <%end if%>
												<br />
  											<%if rsSearchH(13)= true then %>
												<a target="_blank" href="http://www.ocid.uts.edu.au/" title="Chemical Risk Assessment">Chemical Risk Assessment</a>
											<%end if%>
											<br />
											<%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/internal/generalriskmgt.doc" title="General risk assessment form (in Microsoft Word format, 67 Kb).">General Risk Assessment</a> <%end if%></td>
												<td><% Response.Write(rsSearchH(15))%></td>
												<td><% Response.Write(rsSearchH(17))%></td>
												<td><%=dtDate%></td>
											</tr><%
							       end if 
							     end if
							      rsSearchH.MoveNext  
							   wend 
							   end if 
'***********************************************************************************************
	    end select
'*********************************************************************************************** %>
</table>
</div>
</div>
</body>
</html>