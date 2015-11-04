<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%'*****declaring the variables**

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

 'Response.Write(facultyID)
' Response.Write("login : "+loginID)
' Response.Write(cboVal)



'Campbells borrowed code to escape the output 15/6/2006
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
Escape = strReturn
End Function

%>

<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <title>Online Risk Register - QORAs Action Status Report for Deans and Directors</title>
 <%'remove the line below %>
 <script type="text/javascript" src="sorttable.js"></script>
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
</head>

<%
  
  ' Analyse the case of the input and then navigate to the proper function
  
  'case 1 : only faculty
  'case 2 : faculty , login and facility
  
     if facultyID <> 0 and len(loginId) > 1 and cboVal = 0 then
          caseval = 1
      
     elseif  facultyID <> 0 and len(loginId)>1 and cboVal <> 0 then
          caseval = 2
      
     elseif  facultyID <> 0 and len(loginId)<=0 and cboVal = 0 then
          caseval = 3
         
     end if
     
  

  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
  '*********************writting the SQL ******************************
  '------------------------get the faculty for the login ---------------
          select case caseval 
  '************************************************************************************************
          case "3" :  ' Only for faculty selection
          
  '***************************** REPORT FOR ONLY FACULTY SELECTION*********************************     
  
  strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
                          &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
                          &" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
                          &" tblQORA.numCampusId = tblCampus.numCampusID and "_
                          &" tblQORA.strsupervisor = tblfacilitySupervisor.strLoginID and "_
                          &" tblQORA.numfacultyId = "& FacultyId &" "_
                          &" ORDER BY strAssessRisk,dtDateCreated,strRoomName"
  'Response.Write(strSQL)
  set rsSearchH = server.CreateObject("ADODB.Recordset")
  rsSearchH.Open strSQL, Conn, 3, 3     
  strFacultyName = rsSearchH("strFacultyName")     
  
              %>
<body>
<div id="wrapper">
 <div id="content">
  <h1 class="pagetitle">Risk Assessments - Action Status Report</h1> 

<table class="suprlevel">
<tr>
 <th>Faculty/Unit</th>
 <td><%Response.Write(strFacultyName) %></td>
</tr>
</table>
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

                    
'set rsSearchH = server.CreateObject("ADODB.Recordset")
 '                         rsSearchH.Open strSQL, Conn, 3, 3 

if rsSearchH.EOF = TRUE then Response.Write("No records found.")  %>


                        <% if not rsSearchH.EOF then 
                                   %> 
                              <table width="100%" id="table13" class="sortable">
							  <thead>
                                <tr>
                                    <th>Location</th>
                                    <th>Hazardous Task</th>
                                    <th>Hazards</th>
                                    <th>Controls</th>
                                    <th>Risk Level</th>
                                    <th>Further Action Required</th>
                                    
                                    <th>Date Actions Completed</th>
                                    <th>Renewal Date</th>
                                </tr>
								</thead>

								<%
                              while not rsSearchH.EOF 
                                dtDate = dateAdd("yyyy",2,rsSearchH(6))
                                if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
                               ' Response.Write(val)
                               ' Response.Write("Expection caught")
                                Valu = len(rsSearchH(17))
                                 if isnull(valu)   then
                                'Response.Write("Expection caught")
                                 strFacultyName = rsSearchH(19)     
                                 strGivenName = rsSearchH(32)
                                  strSurname = rsSearchH(33)
                                  strName = cstr(strGivenName) + " " + cstr(strSurname)
                                  strLocation = cstr(rsSearchH(25))+"/"+ cstr(rsSearchH(26))
                                  strBuilding = cstr(rsSearchH(36))
                                  strCampus =cstr(rsSearchH(39))
                                      %>
                                  <tbody>
										<tr>
										    <td><% Response.Write(strname)%>,<BR><% Response.Write(strLocation)%>,<BR><%Response.Write(strCampus)%>,<BR><% Response.Write(strBuilding)%></td>
                                            <td><% Response.Write(rsSearchH(8))%></td>
                                            <!--<td><% Response.Write(rsSearchH(11))%></td> -->
											<td><%=Escape(rsSearchH(11))%></td>
                                            <td><% Response.Write(rsSearchH(10))%></td>
                                            <td><%=rsSearchH(9)%></td>
                                            <%'=cstr(rsSearchH(19))+"/"+ cstr(rsSearchH(20))+","+ cstr(rsSearchH(24))+","+ cstr(rsSearchH(27)) %>
                                            <td><%=escape(rsSearchH(15))%><BR><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br />
                                        <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br />    
                                        <%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
                                        Assessment</a> <%end if%></td>
                                         
                                            <td><% Response.Write(rsSearchH(17))%></td>
                                            <td><%=dtDate%></td>
                                        </tr>

                            <%
                                    end if 
                                 end if
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
                        <table class="suprlevel">
						<tr>
							<th colspan="2"><strong><%Response.Write(strFacultyName) %></strong></th>
						</tr>
						<tr>
							<th>Supervisor</th>
							<td><%Response.Write (strName)%></td>
						</tr>
                     <%if not rsSearchfaculty.EOF then%>
						<tr>
							<th>Facility Room Name/Number</th>
							<td><%=cstr(rsSearchFaculty(25))+"/"+ cstr(rsSearchFaculty(26))%></td>
						</tr>
						
						<tr>
							<th>Campus</th>
							<td><%=cstr(rsSearchFaculty(39)) %></td>
						</tr>
						
						<tr>
							<th>Building</th>
							<td><%=cstr(rsSearchFaculty(36)) %></td>
						</tr>

						 </table>
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
                        
                         strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus,tblRiskLevel "_
                              &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
                              &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
                              &" tblQORA.numCampusId = tblCampus.numCampusID and "_
                              &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel and "_
                              &" tblQORA.numFacilityId = "& cboVal &" and "_
                              &" strSupervisor = '"& loginId &"'  ORDER BY tblRiskLevel.numGrade"
                
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
                              <table width="100%" class="sortable suprlevel" id="taz2l">
							  <thead>
                                <tr>
                                    <th>Hazardous Task</th>
                                    <th>Hazards</th>
                                    <th>Controls</th>
                                    <th width="64">Risk Level</th>
                                    <th width="224">Further Action Required</th>
                                    <th>Date Actions Completed</th>
                                    <th>Renewal Date</th>
                                </tr>
                                </thead>
                                <tbody>
								<%


                              while not rsSearchH.EOF 
                                dtDate = dateAdd("yyyy",2,rsSearchH(6))
                                if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
                                 Valu = len(rsSearchH(17))
                                 'Response.Write(val)
                                 'Response.Write("Expection caught")
                                     if isnull(valu)   then
                                           ' Response.Write("Exception caught.")
                                            %>
											
                                            <tr>
                                             <td><!-- HAZARDOUS TASK --><% Response.Write(rsSearchH(8))%></td>
                                             <!--<td><% Response.Write(rsSearchH(11))%></td> ORIGINAL-->
                                             <!-- 15/6/2006 CKL escaped output - puts new Hazards on a new line -->
                                             <td><!-- HAZARDS --><%=Escape(rsSearchH(11))%></td>
											 <!-- 15/6/2006 CKL escaped output - puts new Controls on a new line -->
                                             <!-- <td><% Response.Write(rsSearchH(10))%></td> ORIGINAL-->
											 <td><!-- CONTROLS --><%=Escape(rsSearchH(10))%></td>
                                             <td width="64"><!-- RISK LEVEL --><center><%=rsSearchH(9)%></center></td>
                                             <td width="224"><!-- FURTHER ACTIONS --><%=Escape(rsSearchH(15))%><BR><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br />
                                             <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br />    
                                            <%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk Assessment</a> <%end if%></td>
                                            
                                            <td><!-- DATE ACTIONS COMPLETED --><% Response.Write(rsSearchH(17))%></td>
                                            <td><!-- RENEWAL DATE --><%=dtDate%></td>
                                            </tr><%
                                   end if 
                                 end if 
                                  rsSearchH.MoveNext  
                               wend 
                              end if 
  
  
 case "1" : ' Selection for everything.

'*************************************************************************************************
strSQL = "SELECT * FROM tblQORA,tblFaculty,tblFacility,tblFacilitySupervisor,tblBuilding,tblCampus "_
                          &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
                          &" tblQORA.numBuildingId = tblBuilding.numBuildingID and tblQORA.numFacultyId = tblFaculty.numFacultyID and"_
                          &" tblQORA.numCampusId = tblCampus.numCampusID and "_
                          &" tblQORA.strsupervisor = tblfacilitySupervisor.strLoginID and "_
                          &" tblQORA.numfacultyId = "& FacultyId &" and "_
                          &" strSupervisor = '"& loginId &"'  ORDER BY strAssessRisk,dtDateCreated,strRoomName"
                          
                          set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
                         ' Response.Write(strSQL) 
                          rsSearchFaculty.Open strSQL, Conn, 3, 3     
                          strFacultyName = rsSearchFaculty(19)     
                          strGivenName = rsSearchFaculty(32)
                          strSurname = rsSearchFaculty(33)
                          strName = cstr(strGivenName) + " " + cstr(strSurname)
                          %>
                        <table class="suprlevel">
						<tr>
							<th colspan="2"><strong><%Response.Write(strFacultyName) %></strong></th>
						</tr>
						<tr>
							<th>Supervisor</th>
							<td><%Response.Write (strName)%></td>
						</tr>
                     <%if not rsSearchfaculty.EOF then%>
						<tr>
							<th>Facility Room Name/Number</th>
							<td><%=cstr(rsSearchFaculty(25))+"/"+ cstr(rsSearchFaculty(26))%></td>
						</tr>
						
						<tr>
							<th>Campus</th>
							<td><%=cstr(rsSearchFaculty(39)) %></td>
						</tr>
						
						<tr>
							<th>Building</th>
							<td><%=cstr(rsSearchFaculty(36)) %></td>
						</tr>

						 </table>
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
                        
                         strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus,tblRiskLevel "_
                              &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
                              &" tblQORA.numBuildingId = tblBuilding.numBuildingID and "_
                              &" tblQORA.numCampusId = tblCampus.numCampusID and "_
                              &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel and "_
                              &" tblQORA.numFacilityId = "& cboVal &" and "_
                              &" strSupervisor = '"& loginId &"'  ORDER BY tblRiskLevel.numGrade"
                
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
                              <table width="100%" class="sortable suprlevel" id="taz2l">
							  <thead>
                                <tr>
                                    <th>Hazardous Task</th>
                                    <th>Hazards</th>
                                    <th>Controls</th>
                                    <th width="64">Risk Level</th>
                                    <th width="224">Further Action Required</th>
                                    <th>Date Actions Completed</th>
                                    <th>Renewal Date</th>
                                </tr>
                                </thead>
                                <tbody>
								<%


                              while not rsSearchH.EOF 
                                dtDate = dateAdd("yyyy",2,rsSearchH(6))
                                if rsSearchH(12)= true or  rsSearchH(13)= true or rsSearchH(14)= true  then
                                 Valu = len(rsSearchH(17))
                                 'Response.Write(val)
                                 'Response.Write("Expection caught")
                                     if  isnull(valu)  then
                                            'Response.Write("Exception caught.")
                                            %>
											
                                            <tr>
                                             <td><!-- HAZARDOUS TASK --><% Response.Write(rsSearchH(8))%></td>
                                             <!--<td><% Response.Write(rsSearchH(11))%></td> ORIGINAL-->
                                             <!-- 15/6/2006 CKL escaped output - puts new Hazards on a new line -->
                                             <td><!-- HAZARDS --><%=Escape(rsSearchH(11))%></td>
											 <!-- 15/6/2006 CKL escaped output - puts new Controls on a new line -->
                                             <!-- <td><% Response.Write(rsSearchH(10))%></td> ORIGINAL-->
											 <td><!-- CONTROLS --><%=Escape(rsSearchH(10))%></td>
                                             <td width="64"><!-- RISK LEVEL --><center><%=rsSearchH(9)%></center></td>
                                             <td width="224"><!-- FURTHER ACTIONS --><%=Escape(rsSearchH(15))%><BR><%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><br />
                                             <%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a><%end if%><br />    
                                            <%if rsSearchH(14)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk Assessment</a> <%end if%></td>
                                            
                                            <td><!-- DATE ACTIONS COMPLETED --><% Response.Write(rsSearchH(17))%></td>
                                            <td><!-- RENEWAL DATE --><%=dtDate%></td>
                                            </tr><%
                                   end if 
                                 end if 
                                  rsSearchH.MoveNext  
                               wend 
                              end if 
'***********************************************************************************************************************                               
        end select
'***********************************************************************************************************************                             %>
</tbody>
</table>
</div>
</div>
</body>
</html>