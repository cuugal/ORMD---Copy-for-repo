<%

numFacilityId = request.form("hdnFacilityId")
numFacilityId= cint(numFacilityId)

dim rsShow
dim rsShowHeader
dim date_d
dim date_m
dim date_y
dim dtRdate

%>

<%
numFacilityId = request.form("hdnFacilityId")
'response.write("***"&request.form("hdnFacilityId")&"***")
'response.end
numFacilityId= cint(numFacilityId)
numoperationId = request.Form("operationId")
numoperationId = cint(numoperationId)

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr

strSQL = "SELECT distinct(tblQORA.numQORAId) as numQORAId, tblQORA.numFacultyId, tblQORA.numFacilityId, tblQORA.strSupervisor, "_
 &" strFacultyName, strRoomName,strRoomNumber,tblQORA.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText,strCampusName,strBuildingName, null as strOperationName, strDateActionsCompleted, "_
 &" strGivenName, strSurname, strRiskLevel "_
 
 &" FROM tblFaculty, tblFacility,tblQORA,tblCampus,tblBuilding,tblRiskLevel,tblFacilitySupervisor "_
 
 &" Where tblQORA.numFacilityID = tblFacility.numFacilityID"_
 &" and tblBuilding.numBuildingID = tblFacility.numBuildingID"_
 &" and tblBuilding.numCampusID = tblCampus.numCampusID"_
 
 &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
 &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID"_
 
 &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
 &" and tblFacility.numFacilityID = "&numFacilityID&""_
 
 &" union all "_
 
 &"SELECT distinct(tblQORA.numQORAId) as numQORAId, tblQORA.numFacultyId, tblQORA.numFacilityId, tblQORA.strSupervisor, "_
 &" strFacultyName, null as strRoomName, null as strRoomNumber, tblQORA.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText, null as strCampusName, null as strBuildingName, strOperationName, strDateActionsCompleted, "_
 &" strGivenName, strSurname, strRiskLevel "_
 
 &" FROM tblFaculty, tblOperations ,tblQORA,tblRiskLevel,tblFacilitySupervisor "_
 
 &" Where tblQORA.numOperationID = tblOperations.numOperationId "_
 
 &" and tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID "_
 &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_

 &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
 &" and tblOperations.numOperationID = "&numoperationID&""_
 &" order by strOperationName, numFacilityID "


 set rsShow = Server.CreateObject("ADODB.Recordset")
   'response.write(strSQL)
   'response.end
   rsShow.Open strSQL, conn, 3, 3  

'********************************edit the code up here for checking the    
if rsShow.EOF <> true then 

	tFac = rsShow("numFacultyId")
  	tFaci = rsShow("numFacilityId")
  	first_time = true
%>
<div id="wrapper">

<div id="content">
<table class="suprlevel" style="width: 100%;">
      <caption>
		To edit a risk assessment, click on its title under the &quot;Hazardous Task&quot; heading.
      </caption>
<% 
supName = cstr(rsShow("strGivenName"))+ " " + cstr(rsShow("strSurName")) 

while not rsShow.EOF 

	if tFaci <> rsShow("numFacilityId") or tOper <> rsShow("strOperationName") or first_time then

		if(rsShow("strRoomNumber") <> "") then %>
		  	
		    <tr>    		
		    		<td class="campus">
		    		<strong>Campus: </strong><%=rsShow("strCampusName")%>&nbsp;&nbsp;&nbsp;</td>
		    		<td class="campus" colspan="3"><strong>Building: </strong><%=rsShow("strBuildingName")%>&nbsp;&nbsp;&nbsp;
		    		<strong>Room Name: </strong><%=rsShow("strRoomName")%>&nbsp;&nbsp;&nbsp;</td>
		    		<td class="campus" colspan="3"><strong>Room Number: </strong><%=rsShow("strRoomNumber")%>&nbsp;&nbsp;&nbsp;
		    		</td>
		  		</tr>
		  		<tr>
		  			<td class="campus">
		  			<strong>Supervisor: </strong><%=supName%>&nbsp;&nbsp;&nbsp;</td>
		  			<td class="campus" colspan="6"><strong>Faculty: </strong><%=rsShow("strFacultyName")%>&nbsp;&nbsp;&nbsp;
		  			</td>		
		  		<tr>
		
		<% 
		elseif(rsShow("strOperationName") <> "") then %>
				<tr><td colspan="7">&nbsp</td></tr>
				<tr>
		  			<td class="campus">
		  			<strong>Supervisor: </strong><%=supName%>&nbsp;&nbsp;&nbsp;</td>
		  			<td class="campus" colspan="6"><strong>Operation: </strong><%=rsShow("strOperationName")%>&nbsp;&nbsp;&nbsp;
		  			</td>		
		  		<tr>
			
		<% end if 
		
		
		tFaci = rsShow("numFacilityId")
   		tOper = rsShow("strOperationName")
   		first_time = false
   		%>
   		<tr>
          <th class="haztaskresult">Task</th>
          <th class="assochazards">Hazards</th>
          <th class="currentcontrols">Current Controls</th>
          <th class="risklevel">Risk Level</th>
          <th class="furtheraction">Proposed Controls</th>
          <th class="renewaldate">Renewal Date</th>
          <th class="swms">SWMS</th>
        </tr>
   		<%
	end if

   dtRDate = dateAdd("yyyy",2,rsShow("dtDateCreated"))
%>

        <tr>
          <td><a target="Operation" title="Click to edit this Risk Assessment." href="EditQORA.asp?numCQORAId= <%=rsShow("numQORAId")%> "><%=rsShow("strtaskDescription")%></a></td>
          <!--<td><%=escape(rsShow("strHazardsDesc"))%></td>-->
          <td><%=rsShow("strHazardsDesc")%></td>
          <td><%
          
          testval = rsShow("numQORAId")
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

          <td><center><%=rsShow("strRiskLevel")%></center></td>

            <td><%
          
          testval = rsShow("numQORAId")
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

         <!-- <td><%=rsShow("strDateActionsCompleted")%></td>-->
         <!--  <td><%=dtRDate%></td> -->
         <td><%=rsShow("dtReview")%></td> 
         <td><center>
        <% If rsShow("boolSWMSRequired") = true Then %>
                 <form method="post" action="SWMS.asp">
         <input type="submit" value="SWMS" name="btnSWMS" />
         <input type="hidden" name="hdnQORAId" value="<%=rsShow("numQORAId")%>" />
         <input type="hidden" name="hdnNoSaveBeforeSWMS" value="nosave"/>
         </form>

        <% End if%>
                 </center></td>
        </tr>
        


  <% 
    rsShow.Movenext
  wend
  %>
</table>
</div>

<%else%>
<p>There are currently no Risk Assessments for this facility/operation</p>
<%end if%>