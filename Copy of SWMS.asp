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

	'-------------------------Back to SWMS---------------------------------
	set dcnDb = server.CreateObject("ADODB.Connection")
	dcnDb.Open constr

	set rsSearch = server.CreateObject("ADODB.Recordset")
	'AA Feb 2010 rewrite for correct utilisation of reln QORA:FACILITY:BUILDING:CAMPUS
strSQL = "Select tblQORA.*, tblBuilding.numBuildingID, tblCampus.numCampusID from tblQORA, tblFacility, tblBuilding, tblCampus where numQORAID = "& testval &""_
	&" and tblQORA.numFacilityID = tblFacility.numFacilityID and tblFacility.numBuildingID = tblBuilding.numBuildingID"_
	&" and tblBuilding.numCampusID = tblCampus.numCampusID"
	rsSearch.Open strSQL, dcnDb, 3, 3
	
	numQORAID = rsSearch("numQORAId")
	numCampusId = rsSearch("numCampusID")
	numBuildingId = rsSearch("numBuildingId")
	numFacilityId = rsSearch("numFacilityId")
	numFacultyId = rsSearch("numFacultyId")
	strSuperv = rsSearch("strSupervisor")
	strAssessor = rsSearch("strAssessor")
	dtCreated = rsSearch("dtDateCreated")
	strTaskDescription = rsSearch("strTaskDescription")
	strHazardsDesc = rsSearch("strHazardsDesc")
	strJobSteps = rsSearch("strJobSteps")
	strAssessRisk = rsSearch("strAssessRisk")
	strControlRiskDesc = rsSearch("strControlRiskDesc")
	strText = rsSearch("strText")
	dtDateCreated = rsSearch("dtDateCreated")
	strInherentRisk = rsSearch("strInherentRisk")
	strDate = rsSearch("strDateActionsCompleted")
	boolswms = rsSearch("boolFurtherActionsSWMS")
	boolCRA = rsSearch("boolFurtherActionsChemicalRA")
	boolGRA = rsSearch("boolFurtherActionsGeneralRA")
	strConsultation = rssearch("strConsultation")
	boolSWMSRequired = rssearch("boolSWMSRequired")
	
	
	
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
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
</head>
<body>
<title>EHS Risk Register for Facilities - Edit a Risk Assessment</title>

<div id="wrapper">
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
      		<td align="left"><h2> Task: <%=rsSearch("strTaskDescription")%> </h2></td>
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
         
        <tr>
        <th>Facility</th>
          <td><strong>Campus</strong><br/><%=rsc("strCampusName")%></td>
          <td><strong>Building</strong><br/><%=rsb("strBuildingName")%></td>
          <td><strong>Room Number/Name</strong><br/><%=strFacilityName%></td>
        </tr>
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
          	<td>Review Date:&nbsp;&nbsp;&nbsp;<%=rsSearch("dtReview")%></td>
        </tr>
  </table>
  <br/>
  <strong> Hazards relating to this task</strong>

      <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
      	
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
    </table>

<br/>
  <strong> Safety equipment, training, signage & information</strong>
  
  <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
            <textarea rows = "8" style="width:100%;" name="T3" readonly>
<% 'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& vbCrLf
     		' get the next record
           rsControls.MoveNext
     		wend %>
     		<%=strControlsImplemented%>
</textarea>           
</td>
</tr> 
</table>
<br/>       
<strong> Job steps </strong> <br/>&mdash These "Job Steps" can be edited directly. 
<% 
if isNull(strJobSteps) then
	strJobSteps = "  BEFORE YOU START:"&vbCrLf&vbCrLf&vbCrLf &"  STEPS IN JOB:"&vbCrLf&vbCrLf&vbCrLf&"  WHEN YOU FINISH:"&vbCrLf&vbCrLf&vbCrLf&"  NEVER:"_
    			 &vbCrLf&vbCrLf&vbCrLf&"  ALWAYS:"&vbCrLf&vbCrLf&vbCrLf&"  EMERGENCY PROCEDURES:"&vbCrLf&vbCrLf&vbCrLf
   end if %>
  
  <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
            <textarea rows = "20" style="width:100%;" name="T4" ><%=strJobSteps%>
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