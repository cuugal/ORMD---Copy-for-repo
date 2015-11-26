<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->

<%strLoginId = session("strLoginId")%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">

<%
testval = request.form("hdnQORAID")
NoSaveBeforeSWMS = request.form("hdnNoSaveBeforeSWMS")
'This file will check if a record exists, and then create, or find a record and update,
' so long as the fields in POST are appropriately named.
 %>
 	<!--#INCLUDE FILE="CreateUpdateQORAFromPost.asp"-->
<% 


set dcnDb = server.CreateObject("ADODB.Connection")
dcnDb.Open constr

set rsResults = server.CreateObject("ADODB.Recordset")
'AA Feb 2010 rewrite for correct utilisation of reln QORA:FACILITY:BUILDING:CAMPUS
strSQL = "Select * from tblQORA where numQORAID = "& testval
rsResults.Open strSQL, dcnDb, 3, 3

'response.write strSql
'response.end

if(rsResults("numFacilityId") <> 0) then
	strSQL = "Select tblQORA.*, tblBuilding.numBuildingID, tblCampus.numCampusID, tblFacilitySupervisor.numFacultyID "_
			&"from tblQORA, tblFacility, tblBuilding, tblCampus, tblFacilitySupervisor where numQORAID = "& testval &""_
			&" and tblQORA.numFacilityID = tblFacility.numFacilityID and tblFacility.numBuildingID = tblBuilding.numBuildingID"_
			&" and tblBuilding.numCampusID = tblCampus.numCampusID"_
			&" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"
	set rsSearch = server.CreateObject("ADODB.Recordset")
	rsSearch.Open strSQL, dcnDb, 3, 3
	numCampusId = rsSearch("numCampusID")
	numBuildingId = rsSearch("numBuildingId")
	numFacilityId = rsSearch("numFacilityId")
	numFacultyID = rsSearch("tblFacilitySupervisor.numFacultyID")
	'Response.Write(numFacilityId)
	numOperationId = 0
end if


if(rsResults("numOperationId") <> 0) then
	strSQL = "Select tblQORA.*, tblFacilitySupervisor.numFacultyID"_
	&" from tblQORA, tblOperations, tblFacilitySupervisor where numQORAID = "& testval &""_
	&" and tblQORA.numOperationID = tblOperations.numOperationID"_
	&" and tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID"
	set rsSearch = server.CreateObject("ADODB.Recordset")
	rsSearch.Open strSQL, dcnDb, 3, 3
  numCampusID = 0
  numBuildingId = 0
  numFacilityId = 0
  numOperationID = rsResults("numOperationId")
  numFacultyID = rsSearch("tblFacilitySupervisor.numFacultyID")
end if
	
	numQORAID = rsResults("numQORAId")
		strSuperv = rsResults("strSupervisor")
	strAssessor = rsResults("strAssessor")
	dtCreated = rsResults("dtDateCreated")
	strTaskDescription = rsResults("strTaskDescription")
	strHazardsDesc = rsResults("strHazardsDesc")
	strJobSteps = rsResults("strJobSteps")
	strAssessRisk = rsResults("strAssessRisk")
	strControlRiskDesc = rsResults("strControlRiskDesc")
	strText = rsResults("strText")
	dtDateCreated = rsResults("dtDateCreated")
	strInherentRisk = rsResults("strInherentRisk")
	strDate = rsResults("strDateActionsCompleted")
	boolswms = rsResults("boolFurtherActionsSWMS")
	boolCRA = rsResults("boolFurtherActionsChemicalRA")
	boolGRA = rsResults("boolFurtherActionsGeneralRA")
	strConsultation = rsResults("strConsultation")
	boolSWMSRequired = rsResults("boolSWMSRequired")
	
	set rsSupervisor = server.CreateObject("ADODB.Recordset")
	strSQL = "select strGivenName,strSurName from tblFacilitySupervisor where strLoginId = '"& strSuperv &"'"
	rsSupervisor.Open strSQL, dcnDb, 3, 3
	if not rsSupervisor.EOF then
	strSupervisor = cstr(rsSupervisor("strGivenName")) +" "+cstr(rsSupervisor("strSurname"))
	'response.write(strSupervisor)
	else
	'response.write("no records")
	end if
	
	'------------------------------------------------------------------------------------------------------
	
	set rsF = server.CreateObject("ADODB.Recordset")
	strSQL = "Select * from tblFaculty where numFacultyID = "& numFacultyId 
	rsF.Open strSQL, dcnDb, 3, 3
	'------------------------------------------------------------------------------------------------------
	set rsC = server.CreateObject("ADODB.Recordset")
	strSQL = "Select * from tblCampus where numCampusID = "& numCampusId 
	rsC.Open strSQL, dcnDb, 3, 3
	'------------------------------------------------------------------------------------------------------
	set rsB = server.CreateObject("ADODB.Recordset")
	strSQL = "Select * from tblBuilding where numBuildingID = "& numBuildingId 
	rsB.Open strSQL, dcnDb, 3, 3
	'------------------------------------------------------------------------------------------------------
	set rsFaci = server.CreateObject("ADODB.Recordset")
	strSQL = "Select * from tblFacility where numFacilityID = "& numFacilityId 
	rsFaci.Open strSQL, dcnDb, 3, 3
	'------------------------------------------------------------------------------------------------------
	set rsOper = server.CreateObject("ADODB.Recordset")
	strSQL = "Select * from tblOperations where numOperationID = "& numOperationID 
	rsOper.Open strSQL, dcnDb, 3, 3
	

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
</head>
<body>
<title>EHS Risk Register for Facilities - Edit a Risk Assessment</title>

<div id="wrapperform">
  <div id="content">	
  <table width = 80%>
     	<tr>
      		<td align="left"><h1 class="pagetitle">Safe Work Method Statement (SWMS)</h1> 
   
  </td>
      		<td align="right"> <h2> Risk Assessment Number <%=testval%></h2></td>
      	</tr>
      </table>
      
   <!--<form target="_blank" method="post" action="SWMS-print.asp"> -->
     <form method="post" name="Form1" action="AddSWMS.asp">  
       
  <table width = 80% class="swms">
     	<tr>
      		<td align="left" class="plainbox"><strong>Work Activity:</strong> <%=rsResults("strTaskDescription")%></td>
		</tr>
		<tr>
			
			<td>
			
			<% 'only display save button if the user is logged in.
				If not Trim(Session("strLoginId")) = "" Then %>
					<table width = 80% class="swms">
						<tr align="center" width="50%">
							<td> 
							<input type="submit" value="To Risk Assessment" onclick="Form1.action='EditQORA.asp?numCQORAId=<%=testval%>'; Form1.target='_self'; return true;">
							</td>
							<td >
							<!-- 9April2010 DLJ added target=self -->
							<input type="submit" value="Save SWMS" target="_self" onclick="Form1.action='AddSWMS.asp'; Form1.target='_self'; return true;"/>
							<!--</form>-->
							</td>
						</tr>
					</table>

				<% else %>
				<table width = 80% class="swms">
					<tr align="center" width="50%">
					<td>
					<% if strLoginID = "" then %>
						<input type="button" value="Back to Risk Assessment List" name="Back" onclick="history.back();">
					<% else %>
						<!--input type="button" value="Back to Risk Assessment" name="Back" onclick="location.href(EditQORA.asp)"-->
					<% end if %>
					</td>
					</tr>
				</table>
			  <% End If %>
			</td>


      		<td align="center">
      		<!-- <input type="submit" value="Print preview" /> -->
      		
			<input type="submit" value="Print Preview" target="_blank" onclick="Form1.action='SWMS-print.asp'; Form1.target='_blank'; return true;" />

      		
      		<input type="hidden" name="hdnQORAId" value="<%=testval%>" /> 
	  		<input type="hidden" name="hdnFacilityId" value="<%=numFacilityID%>" />
	  		<input type="hidden" name="operationId" value="<%=numOperationID%>" />
	  		</td>
      		<% if strAssessRisk="L" then%> <td class="low" align="right" width="250"> Residual Risk Level: Low </td><%end if%>
      		<% if strAssessRisk="M" then%> <td class="medium" align="right" width="250"> Residual Risk Level: Medium </td><%end if%>
      		<% if strAssessRisk="H" then%> <td class="high" align="right" width="250"> Residual Risk Level: High </td><%end if%>
      		<% if strAssessRisk="E" then%> <td class="extreme" align="right" width="250"> Residual Risk Level: Extreme </td><%end if%>
      	</tr>
  </table>


  
   <table class="suprreportheader" style="width: 80%; margin-bottom:20px;">

    <tr>
          <th>Faculty/Unit</th>
          <td colspan="3"><%=rsf("strFacultyName")%></td>
        </tr>
<% if(rsResults("numFacilityId") <> 0) then 	
	'code for Facility Name
	set connFacility = Server.CreateObject("ADODB.Connection")
	connFacility.open constr
	' setting up the recordset
	strSQL ="Select * from tblFacility where numFacilityId ="& numFacilityId
	set rsFillFacility = Server.CreateObject("ADODB.Recordset")
	rsFillFacility.Open strSQL, connFacility, 3, 3
	  
	strRoomName = rsFillFacility("strRoomName")
	strRoomNo = rsFillFacility("strRoomNumber")    
	strFacilityName =  cstr(strRoomNo) + " - " + cstr(strRoomName) 
%>
        <tr>
        <th>Facility</th>
          <td><strong>Campus</strong><br/><%=rsc("strCampusName")%></td>
          <td><strong>Building</strong><br/><%=rsb("strBuildingName")%></td>
          <td><strong>Room Number/Name</strong><br/><%=strFacilityName%></td>
        </tr>
<% end if 

if(rsResults("numOperationId") <> 0) then %>
	<tr>
        <th>Operation</th>
        <td colspan="3"><%=rsOper("strOperationName")%></td>
    </tr>
<% end if %>
        <tr>
          <th>Supervisor Name</th>
          <td colspan="3"><%=strSupervisor%></td>
        </tr>
        <tr>
        <%' Code to create an Australian date format
			todaysday = day(date)
			todaysMonth = month(date)
			todaysYear = year(date)
			renewal = todaysYear + 1

			todaysDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(todaysYear)
			renewalDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(renewal)
			%>
        	<th>Assessor:</th><td><%=strAssessor%></td>
        	<td>Date Last Modified:&nbsp;&nbsp;&nbsp;
          	<%=dtDateCreated%></td>
          	<!-- <td>Review Date:&nbsp;&nbsp;&nbsp;<%=renewalDate%></td> -->
          	<td>Date Due for Review:&nbsp;&nbsp;&nbsp;<%=rsResults("dtReview")%></td>
        </tr>
  </table>
  <br/>
  <strong> Hazards relating to this Work Activity</strong>

      <!--table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
		<tr>  
			<td>
				<textarea rows = "8" style="width:100%;" name="T1"readonly >TASK HAZARDS:
				<%=strHazardsDesc%>
				</textarea>           
			</td>
		</tr> 
		<tr>
			<td>
				<textarea rows = "8" style="width:100%;" name="T2" readonly>INHERENT RISKS:
				<%=strInherentRisk%>
				</textarea>
          </td>
		</tr>  
    </table-->


      <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
		<tr>  
			<td>
				<strong>TASK HAZARDS:</strong><br> <%=strHazardsDesc%>          
			</td>
		</tr> 
		<tr>
		<td><br></td>
		</tr>
		<tr>
			<td>
				<strong>INHERENT RISKS:</strong><br>	<%=strInherentRisk%>
          </td>
		</tr>  
    </table>


<br/>
  <strong> Control Measures - Safety equipment, training, signage & information</strong>
  
  <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
            <!--textarea rows = "8" style="width:100%;" name="T3" readonly-->
<% 'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& "<BR>"
     		' get the next record
           rsControls.MoveNext
     		wend %>
     		<%=strControlsImplemented%>
<!--/textarea-->           
</td>
</tr> 
</table>
<br/>       
<strong> Work Activity steps </strong> <br/>&mdash; These "Work Activity Steps" can be edited directly. 
<% 
if isNull(strJobSteps) then
	strJobSteps = "  BEFORE YOU START:"&vbCrLf&"         e.g.inspection or maintenance checks"&vbCrLf&vbCrLf&vbCrLf&"  STEPS IN WORK ACTIVITY (Noting how job is made safe as per the above Control Measures):"&vbCrLf&"  (1)"&vbCrLf&"  (2)"_
    			 &vbCrLf&"  (3) etc ..."&vbCrLf&vbCrLf&vbCrLf&vbCrLf&"  EMERGENCY PROCEDURES:"&vbCrLf&"  Dial 6"&vbCrLf&vbCrLf&vbCrLf&"  Certificates/Licensing/WorkCover Permits Required:"&vbCrLf&vbCrLf&"  Training Required:"&vbCrLf&vbCrLf&"  Codes or Standards Applicable:"&vbCrLf&vbCrLf
   end if %>
  
  <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
            <textarea rows = "30" style="width:100%;" name="T4" ><%=strJobSteps%>
</textarea> 
          
</td>
 </tr> 
</table>


			  <!-- buttons used to live here -->


  </form>
    <div style="clear:both"></div>  
  	</div>
  </div>
  

<br/>
<br/>
  </body>
  </html>