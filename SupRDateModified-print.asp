<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
dim loginId
loginId = session("strLoginId")

QORAtype = session("QORAtype")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!--<link rel="stylesheet" type="text/css" href="orr.css" media="screen,print" />-->
 <link rel="stylesheet" type="text/css" href="orrprint.css" media="screen,print" />
 <title>Online Risk Register - Report for Supervisors</title>
 <script type="text/javascript" src="sorttable.js"></script>
 <style type="text/css">
<!--
.style2 {color: #FFFFFF}
-->
 </style>
</head>
<%
'Campbells borrowed code to escape the output 15/6/2006
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
Escape = strReturn
End Function

'*******declaring the variables****
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
  dim cboValDummy
  dim numOptionId
  dim numPageStatus
  
  

    
      numPageStatus = request.querystring("cboValDummy")
      'Response.Write(numPageStatus) 
'DLJ only need to print one facility location at a time for the facility folder
'     if numPageStatus = "1" then 
'      cboVal = session("cboVal")
    'response.write(cboVal)
'     else
'       cboVal = Request.Form("cboFacility")  
'       session("cboVal")= cboval   
'     end if  
     'response.write(cboVal)
      cboVal = session("cboVal")
      cboOperation = session("cboOperation")


      'Response.Write (Session("cboVal"))    
       numOptionId = Request.QueryString("numOptionID")
      'Response.Write(numOptionId) 
      
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
  strFacultyName = rsSearchFaculty("strFacultyName")     
  strGivenName = rsSearchFaculty("strGivenName")
  strSurname = rsSearchFaculty("strSurname")
  strName = cstr(strGivenName) + " " + cstr(strSurname)
  %>

<body>

<div id="wrapper">

<div id="content">

<!-- outside table -->
<table class="mainprintable">
<tr>
<td>
  <img src="utslogo.gif" width="184" height="41" alt="" align="left" />
</td>
<td>
<h1>Risk Assessment Register</h1>
</td>
<td class="date">Date Printed: <%=date%></td>
</tr>
<tr>
 <td colspan="3">


<!--header table-->
<table class="suprlevel-print">

<%
 '****Writing the report****
 
 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblRiskLevel, tblFacilitySupervisor  "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"' ORDER BY tblRiskLevel.numGrade, strTaskDescription"

 
  'Response.Write(strSQL)
  'insert case here
  'AA jan 2010: altered ALL of the below to remove strFacilitySupervisor	
select case numOptionId
case "1" :

strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblFacilitySupervisor "_
 &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
 &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
 &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
 &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
 &" tblQORA.numFacilityId = "& cboVal &" and "_
 &" strLoginID = '"& loginId &"' ORDER BY strTaskDescription"
case "2" :
				 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblFacilitySupervisor  "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"' ORDER BY strControlRiskDesc"
				
case "3" :
				 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblFacilitySupervisor  "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"'  ORDER BY strRoomName,strRoomNumber,strBuildingName,strCampusName" 

case "4" :
				 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblFacilitySupervisor  "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"' ORDER BY  strAssessRisk" 

case "5" :
				 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblFacilitySupervisor  "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"' ORDER BY strDateActionsCompleted" 

case "6" :
strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus , tblFacilitySupervisor "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  &" tblQORA.numFacilityId = "& cboVal &" and "_
  &" strLoginID = '"& loginId &"'  ORDER BY dtDateCreated"

end select
  
  'Get the login ID out as this is more useful
  strSQL2 = "Select * from tblFacilitySupervisor where strLoginID = '"& loginId &"'" 
  
  set rsID = server.CreateObject("ADODB.Recordset")
  rsID.Open strSQL2, Conn, 3, 3 
  numSupervisorId = rsID("numSupervisorId")
  
if(QORAtype="operation") then 
	strSQL = "SELECT * FROM tblQORA, tblOperations "_
	&" WHERE tblQORA.numOperationID = tblOperations.numOperationID and "_
	&" tblQORA.numOperationID = "&cboOperation&" and "_
	&" numFacilitySupervisorId = "& numSupervisorId &" ORDER BY dtDateCreated" 
	      
end if

set rsSearchH = server.CreateObject("ADODB.Recordset")
rsSearchH.Open strSQL, Conn, 3, 3 

if(QORAtype = "location") then %>
<tr>    		
	<td class="campus">
	<strong>Campus: </strong><%=rsSearchH("strCampusName")%>&nbsp;&nbsp;&nbsp;</td>
	<td class="campus" colspan="3"><strong>Building: </strong><%=rsSearchH("strBuildingName")%>&nbsp;&nbsp;&nbsp;
	<strong>Room Name: </strong><%=rsSearchH("strRoomName")%>&nbsp;&nbsp;&nbsp;</td>
	<td class="campus" colspan="3"><strong>Room Number: </strong><%=rsSearchH("strRoomNumber")%>&nbsp;&nbsp;&nbsp;
	</td>
	</tr>
	<tr>
		<td class="campus">
		<strong>Supervisor: </strong><%=strName%>&nbsp;&nbsp;&nbsp;</td>
		<td class="campus" colspan="6"><strong>Faculty: </strong><%=strFacultyName%>&nbsp;&nbsp;&nbsp;
		</td>		
	</tr>
</table>

<% end if
if(QORAtype = "operation") then %>
		<tr>
  			<td class="campus">
  			<strong>Supervisor: </strong><%=strName%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="6"><strong>Operation: </strong><%=rsSearchH("strOperationName")%>&nbsp;&nbsp;&nbsp;
  			</td>		
  		</tr>
</table>
 <%
end if  
   'Response.Write(strSQL) 
if not rsSearchH.EOF then 
       %> 
	<br />
	<table class="suprlevel-print" id="id13">
	
	<thead>
	<tr>
		 <th>Ra No.</th>
		 <th>Task</th>
		 <th>Hazards</th>
		 <th>Current Controls</th>
		 <th>Risk Level</th>
		 <th>Proposed Controls</th>
		 <th>Review Date</th> 
		</tr>
	</thead>
	<!--caption>Click a table heading to sort by the respective criteria.</caption-->

	<%
	
	 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",1,rsSearchH("strDate"))
    
    %>
 <tbody>

 <tr>
 		<td><%=Escape(rsSearchH("numQORAId"))%></td>
		<td><%=rsSearchH("strTaskDescription")%></td>
<!--		<td><% Response.Write(rsSearchH(11))%></td> -->
		<td><%=Escape(rsSearchH("strHazardsDesc"))%></td>
		<td><%
          
          testval = rsSearchH("numQORAId")
           	'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& "<br/>"
     		' get the next record
           rsControls.MoveNext
     		wend 
     	   %>
     	  
     	<%=strControlsImplemented%>
          
       </td>
		<td><center><%=rsSearchH("strAssessRisk")%></center></td>
		<!--<td><% Response.Write(rsSearchH(15))%><br />
			<%if rsSearchH(12)= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc" title="Safe Work Method Statement (in Microsoft Word format, 47 Kb).">Safe Work Method Statement</a> <%end if%><br /><br />
			<%if rsSearchH(13)= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/" title="Chemical risk assessment at OCID.">Chemical Risk Assessment</a> <%end if%><br /><br />
			<%if rsSearchH(14)= true then %>Detailed Risk Assessment<%end if%></td>
		<td><% Response.Write(rsSearchH(17))%></td>-->
		<td>
		<%
          ' New code to put in the unimplemented risk controls
          
          testval = rsSearchH("numQORAId")
           	'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and not boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& "<br/>"
     		' get the next record
           rsControls.MoveNext
     		wend 
     	   %>    	  
     	<%=strControlsImplemented%>

		</td>
		
		<!-- <td><%=dtDate%></td> -->
		<td><%=rsSearchH("dtReview")%></td>

</tr>



<%
    rsSearchH.MoveNext  
 wend


 
 %>
 </tbody>
 </table>

 <%
 end if 
%>
  

</td>
</tr>
</table>



    <div>
      <h4>Declaration</h4>
      I certify that all hazards have been identified and addressed in this facility.<br />
      <br />
      Signed: ___________________________ <br />
      <br />
      Date: ____________________________ </div>


</div>
</div>
</body>
</html>