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


%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!--<link rel="stylesheet" type="text/css" href="orr.css" media="screen,print" />-->
 <link rel="stylesheet" type="text/css" href="orrprint.css" media="screen,print" />
 <title>Online Risk Register - Report for Supervisors</title>
 <link rel="SHORTCUT ICON" href="favicon.ico" type="image/x-icon" />
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
  dim cboFacility
  dim cboValDummy
  dim numOptionId
  dim numPageStatus
  
  

    
      numPageStatus = request.querystring("cboValDummy")

      cboFacility = session("cboFacility")
      cboOperation = session("cboOperation")
      searchType = session("searchType")

      'Response.Write (Session("cboVal"))    
       numOptionId = Request.QueryString("numOptionID")
      'Response.Write(numOptionId) 
      
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
   

  %>

<body>

<div id="wrapper">

<div id="content">

<!-- outside table -->
<table class="mainprintable">
<tr>
<td>
  			<img src="blackutslogo.png" width="92" height="41" alt="" align="left" />
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

 if(searchType = "location") then 
 'AA jan 2010 rewrite include join to tlFacilitySupervisor as part of reln fix
 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblRiskLevel ,tblFacilitySupervisor, tblFaculty "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  
  &" tblQORA.numFacilityId = "& cboFacility &" and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
   &" tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId and "_
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel  "_ 
  &" ORDER BY tblFacilitySupervisor.numFacultyId, tblRiskLevel.numGrade, strTaskDescription"
 end if
 
 'DLJ July2016 changed numOperationId to cboOperation
 
 if(searchType = "operation") then
	 strSQL = "SELECT * FROM tblQORA, tblOperations, tblRiskLevel ,tblFacilitySupervisor, tblFaculty "_
  &" WHERE tblQORA.numOperationId = tblOperations.numOperationId and "_
  &" tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID and"_
  
  &" tblQORA.numOperationId = "& cboOperation &" and "_
  &" tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId and "_
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel  "_ 
  &" ORDER BY tblFacilitySupervisor.numFacultyId, tblRiskLevel.numGrade, strTaskDescription"
 end if

set rsSearchH = server.CreateObject("ADODB.Recordset")
rsSearchH.Open strSQL, Conn, 3, 3 

if(searchType = "location") then %>
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
		<strong>Supervisor: </strong><%=rsSearchH("strGivenName")&" "&rsSearchH("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
		<td class="campus" colspan="6"><strong>Faculty: </strong><%=rsSearchH("strFacultyName")%></strong>&nbsp;&nbsp;&nbsp;
		</td>		
	</tr>
</table>

<% end if
if(searchType = "operation") then %>
		<tr>
  			<td class="campus">
  			<strong>Supervisor: </strong><%=rsSearchH("strGivenName")&" "&rsSearchH("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
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
		 <th class="hazard-column">Hazards</th>
		 <th class="controls-column">Current Controls</th>
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



    <!--div>
      <h4>Declaration</h4>
      I certify that all hazards have been identified and addressed in this facility.<br />
      <br />
      Signed: ___________________________ <br />
      <br />
      Date: ____________________________ </div-->


</div>
</div>
</body>
</html>