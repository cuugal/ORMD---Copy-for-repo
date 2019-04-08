<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

'Set session to display dates in non-us format
session.LCID = 2057	'English(British) format
%>
<%strLoginId = session("strLoginId")%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">

<%


'***Fetching the details***

testval = request.QueryString("numCQORAID")
testval = cdbl(testval)

'response.write(testval)


set dcnDb = server.CreateObject("ADODB.Connection")
dcnDb.Open constr

set rsResults = server.CreateObject("ADODB.Recordset")
'AA Feb 2010 rewrite for correct utilisation of reln QORA:FACILITY:BUILDING:CAMPUS
strSQL = "Select * from tblQORA where numQORAID = "& testval
rsResults.Open strSQL, dcnDb, 3, 3



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

if(rsResults("numOperationId") = 0 and rsResults("numFacilityId") = 0) then
	strSQL = "Select tblQORA.*, tblFacilitySupervisor.numFacultyID"_
	&" from tblQORA, tblFacilitySupervisor where numQORAID = "& testval &""_
	&" and tblFacilitySupervisor.strLoginId = tblQORA.strSupervisor"
	set rsSearch = server.CreateObject("ADODB.Recordset")
	rsSearch.Open strSQL, dcnDb, 3, 3
  numCampusID = 0
  numBuildingId = 0
  numFacilityId = 0
  numOperationID = 0
  numFacultyId = rsSearch("tblFacilitySupervisor.numFacultyId")
    searchType = "user"
 end if


numQORAID = rsResults("numQORAId")
strSuperv = rsResults("strSupervisor")
strAssessor = rsResults("strAssessor")
dtCreated = rsResults("dtDateCreated")
strTaskDescription = rsResults("strTaskDescription")
strHazardsDesc = rsResults("strHazardsDesc")
strAssessRisk = rsResults("strAssessRisk")
strLikelyhood =  rsResults("strLikelyhood")
strConsequence = rsResults("strConsequence")
strControlRiskDesc = rsResults("strControlRiskDesc")
strText = rsResults("strText")
strInherentRisk = rsResults("strInherentRisk")
strDate = rsResults("strDateActionsCompleted")
boolswms = rsResults("boolFurtherActionsSWMS")
boolCRA = rsResults("boolFurtherActionsChemicalRA")
boolGRA = rsResults("boolFurtherActionsGeneralRA")
strConsultation = rsResults("strConsultation")
boolSWMSRequired = rsResults("boolSWMSRequired")
strJobSteps = rsResults("strJobSteps")

'*****SQL to fetch the name values of the varialbles **********
'----------------------------------------------------------
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

<SCRIPT type="text/javascript" language="Javascript" SRC="validation.js">
</SCRIPT>
<!-- Code for the hover menus -->

<SCRIPT type="text/javascript" language="Javascript" SRC="tabbed.js">
</SCRIPT>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<!--[if IE]>

<link rel="stylesheet" type="text/css" href="IE7.css" />

<![endif]-->
<title>Online Risk Register - Edit an existing Risk Assessment</title>
<link rel="SHORTCUT ICON" href="favicon.ico" type="image/x-icon" />
    <!--#include file="bootstrap.inc"--> 
</head>
<body>
    <!--#include file="HeaderMenu.asp" -->

<form method="post" action="ok.asp" name="Form1" onSubmit="return ConfirmChoice();">
  <input type="hidden" name="hdnBuildingId" value="<%=numBuildingID%>" />
  <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" />
  <input type="hidden" name="hdnfacilityId" value="<%=numFacilityID%>" />
  <input type="hidden" name="hdnFacilityName" value="<%=strFacilityName%>" />
  <input type="hidden" name="hdnQORAID" value="<%=testval%>" />
  <input type="hidden" name="hdnDateCreated" value="<%=dtCreated%>" />
  <input type="hidden" name="hasSWMS" value="<%=strJobSteps%>" />
  <input type="hidden" name="operationId" value="<%=numOperationID%>" />
  <%
Dim connAdmin
Dim rsFillAdmin
Dim connFaculty
Dim rsFillFaculty
Dim numFacultyId

'Database Connectivity Code 
  set connAdmin = Server.CreateObject("ADODB.Connection")
  connAdmin.open constr
  
'Database Connectivity Code 
  set connFaculty = Server.CreateObject("ADODB.Connection")
  connFaculty.open constr
  
' setting up the recordset

strSQL ="Select * from tblFaculty "
  
set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
rsFillFaculty.Open strSQL, connFaculty, 3, 3
Dim strFacultyName   
Dim strSupervisorName
Dim strGivenName
Dim strSurname
   
   %>

  <div id="wrapperform">
  <div id="content">

  <table style="width: 82%">
     	<tr>
      		<td align="left"><h2 class="pagetitle">Edit an existing Risk Assessment</h2></td>
      		<td align="right"> <h2> RA Number <%=numQORAID%></h1></td>
      	</tr>
      </table>
  <table class="suprreportheader" style="width: 82%">
    <input type="hidden" name="VTI-GROUP" value="1" />
 
    <tr>
          <th>Faculty/Unit</th>
          <td colspan="3"><%=rsf("strFacultyName")%></td>
        </tr>
 <%
   if(rsResults("numFacilityId") <> 0) then 
 '*************************** Code to get the details of location************************
 
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

        <%' Code to create an Australian date format
			todaysday = day(date)
			todaysMonth = month(date)
			todaysYear = year(date)
			renewal = todaysYear + 1

			todaysDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(todaysYear)
			renewalDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(renewal)
			%>
        	<th>Assessor/Reviewer</th>


                <td >
                    <input type="text" name="txtAssessor" value="<%=strAssessor%>">
                </td>



        	<td>Date Last Modified &nbsp;&nbsp;&nbsp;
          	<!--input type="text" name="txtDateCreated" size="9" value="<%=todaysDate%>"/></td-->
			<%=todaysDate%></td>
          	<!--<td>Review Date&nbsp;&nbsp;&nbsp;<%=renewalDate%></td>-->
          	<td>Current Review Date&nbsp;&nbsp;&nbsp;<%=rsResults("dtReview")%></td>
        </tr>
		<tr>
			<!--DLJ dummy placeholder text box-->
			<th>Persons Consulted</th>
			<td colspan="3"><textarea rows="1" name="strConsultation" cols="90" ><%=strConsultation%></textarea></td>
		</tr>
  </table>

    </br>
	  <table class="reportwarn" style="width: 82%">
      <tr><td><b>NOTE: All risk assessments should be performed in consultation with staff involved with this work activity.</b></td></tr>
	  </table>
	</br>


	      	<strong>(1) Work activity description</strong><br>
      <br>
      <table class="suprreportheader" style="width: 82%">
        <tr>
		  <td width = 30%><b>Briefly describe this hazardous work activity:</b> </br>E.g. Operating ..., Handling ..., Using ...</br>Include names of hazardous equipment, substances or materials used, </br> and any quantities and concentrations of substance(s) or reaction products</td>
          <td><!--<input type="text" name="txtTaskDesc" size="100%" />-->
          <textarea rows="4" name="txtTaskDesc" cols="100" ><%=strTaskDescription%></textarea></td>
        </tr>
      </table>
  
  
  <BR>
      <hr style = "width: 82%;" align="left" />
  
  <!--Navigation -->
  <a name="point1"></a> <strong>(2) Select hazards relating to work activity</strong>
  <p style = "width:80%">Select from the menu below all of the hazards that apply to the work activity.<br />  NOTE: Lists of hazards appear when you put the cursor over this menu. When you click on one it appears in the text box below.</p>
<div>
 <table class="suprreportheader" style="width: 82%">
 	<tr>
		<td colspan="2" >



<!-- start new style tabs grid -->
<ul class="nav nav-tabs" >
   <li class="active"><a data-toggle="tab" href="#environment">Working Environment</a></li>
   <li><a data-toggle="tab" href="#ergonomic">Ergonomic /Manual Handling</a></li>
   <li><a data-toggle="tab" href="#plant">Plant</a></li>
   <li><a data-toggle="tab" href="#electrical">Electrical</a></li>
   <li><a data-toggle="tab" href="#chemical">Chemical</a></li>
   <li><a data-toggle="tab" href="#biological">Biological</a></li>
   <li><a data-toggle="tab" href="#radiation">Radiation</a></li>
</ul>
<style type="text/css">
   .nav::after{
   display:block !important;
   }
   .tab-content ul{
   list-style: outside none none;
   }
   .tab-content ul li{
   padding:5px 0 5px 0;
   }
   .tab-content ul li a:link, .tab-content ul li a:visited {
   background: transparent url("images/link_dot.gif") no-repeat scroll left top;
   color: #000;
   padding-left: 15px;
   }
</style>
<div class="tab-content">
   <div id="environment" class="tab-pane fade in active">
      <div class="row">
         <div class="col-xs-4">
            <ul>
               <li><a href="#point1" onClick="Populate('Working Environment - Working in Remote Locations\r\n')" title="Click to add 'Working in Remote Locations' as a Hazard in this Risk Assessment.">Working in Remote Locations</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Working Outdoors\r\n')" title="Click to add 'Working Outdoors' as a Hazard in this Risk Assessment.">Working Outdoors</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Clinical/Industrial Placements\r\n')" title="Click to add 'Clinical/Industrial Placements' as a Hazard in this Risk Assessment.">Clinical/Industrial Placements</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Violent or Volatile Clients/Interviewees\r\n')" title="Click to add 'Violent or Volatile Clients/Interviewees' as a Hazard in this Risk Assessment.">Violent or Volatile Clients/Interviewees</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Poor Ventilation/Air Quality\r\n')" title="Click to add 'Poor Ventilation/Air Quality' as a Hazard in this Risk Assessment.">Poor Ventilation/Air Quality</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul >
               <li><a href="#point1" onClick="Populate('Working Environment - Work Area Not Suited to Task\r\n')" title="Click to add 'Work Area Not Suited to Task' as a Hazard in this Risk Assessment.">Work Area Not Suited to Task</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Extremes in Temperature\r\n')" title="Click to add 'Extremes in Temperature' as a Hazard in this Risk Assessment.">Extremes in Temperature</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Confined Space\r\n')" title="Click to add 'Confined Space' as a Hazard in this Risk Assessment.">Confined Space</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Isolation\r\n')" title="Click to add 'Isolation' as a Hazard in this Risk Assessment.">Isolation</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point1" onClick="Populate('Working Environment - Fieldwork\r\n')" title="Click to add 'Fieldwork' as a Hazard in this Risk Assessment.">Fieldwork</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Working at a Height\r\n')" title="Click to add 'Working at a Height' as a Hazard in this Risk Assessment.">Working at a Height</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Slip and Trip Hazards\r\n')" title="Click to add 'Dangerous Goods' as a Hazard in this Risk Assessment.">Slip and Trip Hazards</a></li>
               <li><a href="#point1" onClick="Populate('Working Environment - Dangerous Goods\r\n')" title="Click to add 'Dangerous Goods' as a Hazard in this Risk Assessment.">Dangerous Goods</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="ergonomic" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Repetitive or Awkward Movements\r\n')" title="Click to add 'Repetitive or Awkward Movements' as a Hazard in this Risk Assessment.">Repetitive or Awkward Movements</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Lifting Heavy Objects\r\n')" title="Click to add 'Lifting Heavy Objects' as a Hazard in this Risk Assessment.">Lifting Heavy Objects</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Lifting Awkward Objects\r\n')" title="Click to add 'Lifting Awkward Objects' as a Hazard in this Risk Assessment.">Lifting Awkward Objects</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Overreaching\r\n')" title="Click to add 'Overreaching' as a Hazard in this Risk Assessment.">Overreaching</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Working Above Shoulder Height\r\n')" title="Click to add 'Working Above Shoulder Height' as a Hazard in this Risk Assessment.">Working Above Shoulder Height</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Working Below Knee Height\r\n')" title="Click to add 'Working Below Knee Height' as a Hazard in this Risk Assessment.">Working Below Knee Height</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Handling Hot Items\r\n')" title="Click to add 'Handling Hot Items' as a Hazard in this Risk Assessment.">Handling Hot Items</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Cramped/Awkward Positioning\r\n')" title="Click to add 'Cramped/Awkward Positioning' as a Hazard in this Risk Assessment.">Cramped / Awkward Positioning</a></li>
               <li><a href="#point1" onClick="Populate('Ergonomic/Manual Handling - Desktop/Bench Height Inappropriate\r\n')" title="Click to add 'Desktop/Bench Height Inappropriate' as a Hazard in this Risk Assessment.">Desktop / Bench Height Inappropriate</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="plant" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Plant - Noise\r\n')" title="Click to add 'Noise' as a Hazard in this Risk Assessment.">Noise</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Vibration\r\n')" title="Click to add 'Vibration' as a Hazard in this Risk Assessment.">Vibration</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Compressed Gas\r\n')" title="Click to add 'Compressed Gas' as a Hazard in this Risk Assessment.">Compressed Gas</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Lifts\r\n')" title="Click to add 'Lifts' as a Hazard in this Risk Assessment.">Lifts</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Hoists\r\n')" title="Click to add 'Hoists' as a Hazard in this Risk Assessment.">Hoists</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point1" onClick="Populate('Plant - Cranes\r\n')" title="Click to add 'Cranes' as a Hazard in this Risk Assessment.">Cranes</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Sharps/Needles\r\n')" title="Click to add 'Sharps/Needles' as a Hazard in this Risk Assessment.">Sharps/Needles</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Moving Parts (Crushing, Friction, Stab, Cut, Shear)\r\n')" title="Click to add 'Moving Parts (Crushing, Friction, Stab, Cut, Shear)' as a Hazard in this Risk Assessment.">Moving Parts (Crushing, Friction, Stab, Cut, Shear)</a></li>
               <li><a href="#point1" onClick="Populate('Plant - Pressure Vessels and Boilers\r\n')" title="Click to add 'Pressure Vessels and Boilers' as a Hazard in this Risk Assessment.">Pressure Vessels and Boilers</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="electrical" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Electrical - Plug-In Equipment\r\n')" title="Click to add 'Plug-In Equipment' as a Hazard in this Risk Assessment.">Plug-In Equipment</a></li>
               <li><a href="#point1" onClick="Populate('Electrical - High Voltage\r\n')" title="Click to add 'High Voltage' as a Hazard in this Risk Assessment.">High Voltage</a></li>
               <li><a href="#point1" onClick="Populate('Electrical - Exposed Conductors\r\n')" title="Click to add 'Exposed Conductors' as a Hazard in this Risk Assessment.">Exposed Conductors</a></li>
               <li><a href="#point1" onClick="Populate('Electrical - Electrical Wiring\r\n')" title="Click to add 'Electrical Wiring' as a Hazard in this Risk Assessment.">Electrical Wiring</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="chemical" class="tab-pane fade">
		<div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
			   <li><a href="#point1" onClick="Populate('Chemical - Explosive (H200-205)\r\n'); " title="Click to add 'Explosive' as a Hazard in this Risk Assessment.">Explosive (H200-205)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Flammable gas (H220)\r\n'); " title="Click to add 'Flammable gas' as a Hazard in this Risk Assessment.">Flammable gas (H220)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Oxidising Gas (H270)\r\n'); " title="Click to add 'Oxidising Gas' as a Hazard in this Risk Assessment.">Oxidising Gas (H270)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Gas under Pressure (H280-281)\r\n'); " title="Click to add 'Gas under Pressure' as a Hazard in this Risk Assessment.">Gas under Pressure (H280-281)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Flammable liquid (H224-227)\r\n');" title="Click to add 'Flammable liquid' as a Hazard in this Risk Assessment.">Flammable Liquid (H224-227)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Flammable Solid (H228)\r\n'); " title="Click to add 'Flammable Solid' as a Hazard in this Risk Assessment.">Flammable Solid (H228)</a></li>
				<li><a href="#point1" onClick="Populate('Chemical - Self-reactive substance (H240-242)\r\n'); " title="Click to add 'Self-reactive substance' as a Hazard in this Risk Assessment.">Self-reactive substance (H240-242)</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul class="section-list">               
               <li><a href="#point1" onClick="Populate('Chemical - Pyrophoric Substance (H250)\r\n'); " title="Click to add 'Pyrophoric Substance' as a Hazard in this Risk Assessment.">Pyrophoric Substance (H250)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Oxidising Solid  (H271-272)\r\n'); " title="Click to add 'Oxidising Solid' as a Hazard in this Risk Assessment.">Oxidising Solid (H271-272)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Oxidising liquid (H271-272)\r\n'); " title="Click to add 'Oxidising liquid' as a Hazard in this Risk Assessment.">Oxidising Liquid (H271-272)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Dangerous when wet (H260-261)\r\n'); " title="Click to add 'Dangerous when Wet' as a Hazard in this Risk Assessment.">Dangerous when wet (H260-261)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Organic Peroxide (H240-242)\r\n'); " title="Click to add 'Organic Peroxide' as a Hazard in this Risk Assessment.">Organic Peroxide (H240-242)</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Corrosive to metals (H290)\r\n'); " title="Click to add 'Corrosive to metals' as a Hazard in this Risk Assessment.">Corrosive to metals (H290)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Skin corrosion /irritation (H314-315)\r\n'); " title="Click to add 'Skin corrosion /irritation' as a Hazard in this Risk Assessment.">Skin corrosion /irritation (H314-315)</a></li>
            </ul>
         </div>
		 <div class="col-xs-4">
            <ul class="section-list">
				<li><a href="#point1" onClick="Populate('Chemical - Toxic  (H300-302, H310-312, H330-332)\r\n'); " title="Click to add 'Toxic' as a Hazard in this Risk Assessment.">Acute Toxicity (H300-302, H310-312, H330-332)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Serious eye damage / eye irritation (H318-319)\r\n'); " title="Click to add 'Skin/Eye Irritant' as a Hazard in this Risk Assessment.">Serious eye damage / eye irritation (H318-319)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Sensitiser (H334, H317)\r\n'); " title="Click to add 'Sensitiser' as a Hazard in this Risk Assessment.">Sensitiser (H334, H317)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Mutagen (H340-341)\r\n'); " title="Click to add 'Mutagen' as a Hazard in this Risk Assessment.">Mutagen (H340-341)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Carcinogen (H350-351)\r\n'); " title="Click to add 'Carcinogen' as a Hazard in this Risk Assessment.">Carcinogen (H350-351)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Toxic to reproduction (H360-362)\r\n'); " title="Click to add 'Toxic to reproduction' as a Hazard in this Risk Assessment.">Toxic to reproduction (H360-362)</a></li>
			   <li><a href="#point1" onClick="Populate('Chemical - Aquatic toxicity\r\n'); " title="Click to add 'Aquatic toxicity' as a Hazard in this Risk Assessment.">Aquatic toxicity</a></li>
            </ul>
         </div>
		</div>
	   </div>
	   <div id="biological" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Biological - Imported Biomaterials\r\n')" title="Click to add 'Imported Biomaterials' as a Hazard in this Risk Assessment.">Imported Biomaterials</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Pathogens\r\n')" title="Click to add 'Pathogens' as a Hazard in this Risk Assessment.">Pathogens</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Infectious Materials\r\n')" title="Click to add 'Infectious Materials' as a Hazard in this Risk Assessment.">Infectious Materials</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Biological - Blood/Bodily Fluids\r\n')" title="Click to add 'Blood/Bodily Fluids' as a Hazard in this Risk Assessment.">Blood/Bodily Fluids</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Genetically Modified Organisms\r\n')" title="Click to add 'Genetically Modified Organisms' as a Hazard in this Risk Assessment.">Genetically Modified Organisms</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Communicable Diseases\r\n')" title="Click to add 'Communicable Diseases' as a Hazard in this Risk Assessment.">Communicable Diseases</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Animal bites and scratches\r\n')" title="Click to add 'Animal bites and scratches' as a Hazard in this Risk Assessment.">Animal bites and scratches</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Biological - Allergies to Animal Bedding, Dander and Fluids\r\n')" title="Click to add 'Allergies to Animal Bedding, Dander and Fluids' as a Hazard in this Risk Assessment.">Allergies to Animal Bedding, Dander and Fluids</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Working with Insects\r\n')" title="Click to add 'Working with Insects' as a Hazard in this Risk Assessment.">Working with Insects</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Working with Fungi/Bacteria/Viruses\r\n')" title="Click to add 'Working with Fungi/Bacteria/Viruses' as a Hazard in this Risk Assessment.">Working with Fungi/Bacteria/Viruses</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="radiation" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Radiation - Ionising Radiation Sources/Equipment\r\n')" title="Click to add 'Ionising Radiation Sources/Equipment' as a Hazard in this Risk Assessment.">Ionising Radiation Sources/Equipment</a></li>
               <li><a href="#point1" onClick="Populate('Radiation - Non-Ionising Radiation (Lasers, Microwaves, Ultraviolet Light)\r\n')" title="Click to add 'Non-Ionising Radiation (Lasers, Microwaves, Ultraviolet Light)' as a Hazard in this Risk Assessment.">Non-Ionising Radiation (Lasers, Microwaves, Ultraviolet Light)</a></li>
            </ul>
         </div>
      </div>
   </div>
</div>
<!---- end new style grid -->



		</td>
	</tr>

	<tr>
    	<td><strong>Hazard List</strong> <br/> List hazards below either by using the  menu above</br> or typing directly into the text box.</td>
      	<td><strong>Potential Harm</strong> <br/>List <strong>injury/illness</strong> that could occur from this work activity and <strong> <u>how</u></strong> this injury/illness could happen.</br> <i>e.g. - Back strain may occur from incorrect lifting, chemical burn via skin contact, chronic illness due to inhalation of fumes, infection due to aerosol inhalation</i></td>

	</tr>
  	<tr>  
    	<td><!-- textarea box goes in this table cell -->
        <br/>
        	<textarea rows="8" name="T1" id="T1" cols="55"><%=strHazardsDesc%></textarea>
            
        </td>

<!-- insert Exposure text area here -->

        <td><!-- textarea box goes in this table cell -->
        	<br />
          	<textarea rows="8" name="T3" id="T3" cols="65" ><%=strInherentRisk%></textarea>
		</td>
	</tr>
</table>


	</br>
	<table class="reportwarn" style="width: 82%">
		<tr>
			<td colspan="2"> Do the hazards you have selected have the potential to cause death, or serious injury / illness (causing temporary disability)? &nbsp 
			<strong>YES</strong> 
			<input type="radio" name="boolSWMSRequired" Value="Yes" onClick="return radiounClick (event)"<%if boolSWMSRequired then%>checked<%end if%> />  &nbsp &nbsp 
			<strong>NO</strong> 
			<input type="radio" name="boolSWMSRequired" Value="No" onClick="return radiounClick (event)" <%if not boolSWMSRequired then%>checked<%end if%> /><br />
			If they do, then a Safe Work Method Statement (SWMS) must also be recorded.	
			</td>
		</tr>
	</table>


       
		<!-- </table>  -->
	</table>
 </div>
    <div style="clear:both">
    </div> 
   
  <BR>
      <hr style = "width: 82%;" align="left" />
  <strong>(3) Select safety control measures to make work activity safe</strong>
      <p style = "width:82%">Select the safety control measures needed to minimise the risk of harm. List the Safety Control Measures that are both 'currently in place' and 'proposed'. NOTE: Lists of safety control measures appear when you put the cursor over this menu.</p>
      
<div>

	<!-- <table class="suprreportheader" style="width: 80%"> -->
	<table class="suprreportheader" name="controls" id="tblControls" style="width: 82%">
	<tr>
		<td colspan="4">




<!-- start new style tabs -->
<ul class="nav nav-tabs" >
   <li class="active"><a data-toggle="tab" href="#eliminate">Eliminate / Isolate / Substitute <br/>/ Engineering controls</a></li>
   <li><a data-toggle="tab" href="#assess">Admin. Specific: Assessments <br/>/ Licences / Work Methods</a></li>
   <li><a data-toggle="tab" href="#ppe">Personal Protective Equipment (PPE)<br/>&nbsp;</a></li>
   <li><a data-toggle="tab" href="#emergency">Emergency Response Systems<br/>&nbsp;</a></li>
</ul>
<div class="tab-content">

   <div id="eliminate" class="tab-pane fade in active">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point2" onClick="PopulateNext('- Remove Hazard\r\n')" title="Add the control 'Remove Hazard' to this Risk Assessment.">Remove Hazard</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Restricted Access\r\n')" title="Add the control 'Restricted Access' to this Risk Assessment.">Restricted Access</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Use safer materials or chemicals\r\n')" title="Add the control 'Use safer materials or chemicals' to this Risk Assessment.">Use safer materials or chemicals</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Redesign the equipment\r\n')" title="Add the control 'Redesign the equipment' to this Risk Assessment.">Redesign the equipment</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Guarding/Barriers\r\n')" title="Add the control 'Guarding/Barriers' to this Risk Assessment.">Guarding / Barriers</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Biosafety Cabinet\r\n')" title="Add the control 'Biosafety Cabinet' to this Risk Assessment.">BioSafety Cabinet</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Fume Cupboard\r\n')" title="Add the control 'Fume Cupboard' to this Risk Assessment.">Fume Cupboard</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Local Exhaust Ventilation\r\n')" title="Add the control 'Local Exhaust Ventilation' to this Risk Assessment.">Local Exhaust Ventilation</a></li>
            </ul>
         </div>
		 <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- General Ventilation\r\n')" title="Add the control 'General Ventilation' to this Risk Assessment.">General Ventilation</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Regular Maintenance of Equipment\r\n')" title="Add the control 'Regular Maintenance of Equipment' to this Risk Assessment.">Regular Maintenance of Equipment</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Redesign the Workspace / Workflow\r\n')" title="Add the control 'Redesign the Workspace / Workflow' to this Risk Assessment.">Redesign the Workspace / Workflow</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Lifting Equipment/Trolleys\r\n')" title="Add the control 'Lifting Equipment/Trolleys' to this Risk Assessment.">Lifting Equipment / Trolleys</a></li>
            </ul>
         </div>
      </div>
   </div>

   <div id="assess" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Training/Information/Instruction\r\n')" title="Click to add the administrative control 'Training/Information/Instruction' to this Risk Assessment.">Training / Information / Instruction</a></li>
               <!--li><a href="#point2" onClick="PopulateNext('- SWMS (Safe Work Method Statement)\r\n')" title="Click to add the administrative control 'SWMS (Safe Work Method Statement)' to this Risk Assessment.">SWMS (Safe Work Method Statement)</a></li-->
               <li><a href="#point2" onClick="PopulateNext('- Document Chemical Risk Assessment in OCID\r\n')" title="Click to add the administrative control 'Chemical Risk Assessment' to this Risk Assessment.">Document Chemical Risk Assessment in OCID</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Licensing/Certification of Operators\r\n')" title="Click to add the administrative control 'Licensing/Certification of Operators' to this Risk Assessment.">Licensing/Certification of Operators</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Test and Tag Electrical Equipment\r\n')" title="Click to add the administrative control 'Test and Tag Electrical Equipment' to this Risk Assessment.">Test and Tag Electrical Equipment</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Monitor Exposure Level (Sound/Substance/Radiation)\r\n')" title="Click to add the administrative control 'Monitor Exposure Level (Sound/Substance/Radiation)' to this Risk Assessment.">Monitor Exposure Level</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Licences (Lifts, Boilers, Pressure Vessles, Radiation)\r\n')" title="Click to add the administrative control 'Licences (Lifts, Boilers, Pressure Vessles, Radiation)' to this Risk Assessment.">Licences (Lifts, Boilers, Pressure Vessels, Radiation)</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Restricted Access\r\n')" title="Click to add the administrative control 'Restricted Access' to this Risk Assessment.">Restricted Access</a></li>               
			   <li><a href="#point2" onClick="PopulateNext('- Regular Breaks\r\n')" title="Click to add the administrative control 'Regular Breaks' to this Risk Assessment.">Regular Breaks</a></li>
             </ul>
         </div>
         <div class="col-xs-4">
            <ul>
			   <li><a href="#point2" onClick="PopulateNext('- Task Rotation\r\n')" title="Click to add the administrative control 'Task Rotation' to this Risk Assessment.">Task Rotation</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Supervision\r\n'); " title="Click to add the administrative control 'Supervision' to this Risk Assessment.">Supervision</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Ladder/Sling Register\r\n'); " title="Click to add the administrative control 'Ladder/Sling Register' to this Risk Assessment.">Ladder / Sling Register</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Work in Pairs\r\n')" title="Click to add the administrative control 'Work in Pairs' to this Risk Assessment.">Work in Pairs</a></li>
            </ul>
         </div>
      </div>
   </div>

   <div id="ppe" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul >
               <li><a href="#point2" onClick="PopulateNext('-  ... type gloves\r\n')" title="Click to add the risk control 'Gloves' to this Risk Assessment.">Gloves (appropriate type)</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Safety Footwear\r\n')" title="Click to add the risk control 'Safety Footwear' to this Risk Assessment.">Safety Footwear</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Safety Glasses/Goggles\r\n')" title="Click to add the risk control 'Safety Glasses/Goggles' to this Risk Assessment.">Safety Glasses / Goggles</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Face Shield\r\n')" title="Click to add the risk control 'Face Shield' to this Risk Assessment.">Face Shield</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Hard Hat\r\n')" title="Click to add the risk control 'Hard Hat' to this Risk Assessment.">Hard Hat</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Respirator/Dust Mask\r\n')" title="Click to add the risk control 'Respirator/Dust Mask' to this Risk Assessment.">Respirator / Dust Mask</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Hearing Protection\r\n')" title="Click to add the risk control 'Hearing Protection' to this Risk Assessment.">Hearing Protection</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Protective Clothing / Lab Coat / Overalls\r\n')" title="Click to add the risk control 'Protective Clothing/Lab Coat/Overalls' to this Risk Assessment.">Protective Clothing / Lab Coat / Overalls</a></li>
            </ul>
         </div>
      </div>
   </div>


	<div id="emergency" class="tab-pane fade">
		  <div class="row">
			 <div class="col-xs-4">
				<ul>
				   <li><a href="#point2" onClick="PopulateNext('- First Aid Kit\r\n')" title="Click to add the risk control 'First Aid Kit' to this Risk Assessment.">First Aid Kit</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Chemical Spill Kit\r\n')" title="Click to add the risk control 'Chemical Spill Kit' to this Risk Assessment.">Chemical Spill Kit</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Extended First Aid Kit\r\n')" title="Click to add the risk control 'Extended First Aid Kit' to this Risk Assessment.">Extended First Aid Kit</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Evacuation/Fire Control\r\n')" title="Click to add the risk control 'Evacuation/Fire Control' to this Risk Assessment.">Evacuation / Fire Control</a></li>
				</ul>
			 </div>
			 <div class="col-xs-4">
				<ul>
				   <li><a href="#point2" onClick="PopulateNext('- Safety Shower\r\n')" title="Click to add the risk control 'Safety Shower' to this Risk Assessment.">Safety Shower</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Eye Wash Station\r\n')" title="Click to add the risk control 'Eye Wash Station' to this Risk Assessment.">Eye Wash Station</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Emergency Stop Button\r\n')" title="Click to add the risk control 'Emergency Stop Button' to this Risk Assessment.">Emergency Stop Button</a></li>
				   <li><a href="#point2" onClick="PopulateNext('- Remote Communication Mechanism\r\n')" title="Click to add the risk control 'Remote Communication Mechanism' to this Risk Assessment.">Remote Communication Mechanism</a></li>
				</ul>
			 </div>
		  </div>
	   </div>

</div>
<!-- end new style tabs -->



		</td>
	</tr>
	

	<tr>
		<th colspan="4" style="text-align:center; min-width:300px;"><strong>Safety Control Measures</th>
	</tr>
	<tr>
		<th colspan="1" style="text-align:center"><strong>Currently in Place and Proposed</strong></th>
		<th colspan="2" style="text-align:center"><strong>Currently In Place &nbsp &nbsp OR &nbsp &nbsp  Proposed Implementation Date (dd/mm/yyyy)</strong></th>
		<th style="text-align:center"><strong>Remove</strong></th>
	</tr>
    <tr>
		<td colspan="1" style="font-size: 7pt;text-align:center">If you cannot find the desired control in the menu,<br/>you can add your own by clicking 'Add Row'.</td>
       	<td style="font-size: 7pt;text-align:center">Tick the checkbox when the safety control measure is in place.</td>
		<td style="font-size: 7pt;text-align:center">You must enter proposed implementation date for proposed safety control measures.</td>
       	<td style="font-size: 7pt;text-align:center">Click remove to delete this row.</td></tr>       		
	</tr>
        	
       	
        	<% 'here we need to populate the table with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	'The top record is the menu, then header, then desc, so start at 3
        	i=3
        	while not rsControls.EOF %>
         	<tr>
          		<td colspan="1" ><input type="text" id="<%="txtRow"&i%>" name="<%="txtRow"&i%>" size="65" value="<%=rsControls("strControlMeasures")%>" /></td>
          		<td align="center"><input type="checkbox" id="<%="selRow"&i%>" name = "<%="selRow"&i%>"<%if rsControls("boolImplemented") then%>checked <%end if%> 
          		onclick="disableProposed(<%=i%>)" /></td>
				<!--DLJ just put this textbox in as a dummy placeholder -->
          		<td align="center"><input size="9" type="text" id="<%="dateRow"&i%>" name = "<%="dateRow"&i%>" value="<%=rsControls("dtProposed")%>" 
          		onblur="isDate(value)" <%if rsControls("boolImplemented") then%>disabled <%end if%> /> </td>
          		<td align="center"><input type="checkbox" /></td>
         	</tr>
     		<% ' get the next record
     		i= i+1
           rsControls.MoveNext
     		wend %>
     		
		</table>

       	<table class="bluebox" style="width:82%;">
       		<tfoot><tr><td style="text-align:right"></td>
       			<!--<input type="button" value="Remove" onclick="removeRowFromTable();" />-->
       			
       			<td style="min-width:240px" colspan="1">&nbsp;&nbsp;&nbsp;</td>
       			
       			<td style="text-align:right">
       			<input type="button" value="Add Row" onclick="addRowToTable('', false);" style="margin-right:85px" />
       			<input type="button" value="Remove" onclick="deleteRow();" style="margin-right:35px" /></td>
       		</tr></tfoot>
		</table>

      <div style="clear:both">
      </div> 
   </div>
  
  
      <hr style = "width: 82%;" align="left" />

<strong>(4) Assess level of residual risk</strong>
  <p style = "width:82%">Use the risk matrix below as a guide to assess the level of risk, based on the hazards identified above and the way that the work is done with safety control measures that are in place.</p>
<!-- Risk Matrix -->

<TABLE>
<TR>
<TD><strong>L<br>I<br>K<br>E<br>L<br>I<br>H<br>O<br>O<br>D</strong></TD>
<TD>
	<TABLE class="eq" >
	<TR>
	<td colspan="6">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
	<strong>CONSEQUENCE</strong></td>
	</TR>

	<TR>
		<TD>&nbsp;</TD>
		<td></td>
		<TD class="hed"><input title="Assess this as insignificant" type="radio" name="radioc" value="Insignificant"   <%if strConsequence = "Insignificant" then%>checked<%end if%> />Insignificant</TD>
		<TD class="hedalt"><input title="Assess this as minor" type="radio" name="radioc" value="Minor"     <%if strConsequence = "Minor" then%>checked<%end if%> />Minor</TD>
		<TD class="hed"><input title="Assess this as moderate" type="radio" name="radioc" value="Moderate"  <%if strConsequence = "Moderate" then%>checked<%end if%> />Moderate</TD>
		<TD class="hedalt"><input title="Assess this as major" type="radio" name="radioc" value="Major"    <%if strConsequence = "Major" then%>checked<%end if%> />Major</TD>
		<TD class="hed"><input title="Assess this as catastrophic" type="radio" name="radioc" value="Catastrophic"  <%if strConsequence = "Catastrophic" then%>checked<%end if%>/>Catastrophic</TD>
	</TR>
	<tr>
		<td></td>
		<td class= "hedalt" style="font-size: 7pt"><strong>Injury/illness<br/>consequence</strong></td>
		<td class= "hed" style="font-size: 7pt">Non-injury incident</td>
		<td class= "hedalt" style="font-size: 7pt">Injury/ill health<br/>requiring first aid</td>
		<td class= "hed" style="font-size: 7pt">Injury/ill health medical<br/>attention</td>
		<td class= "hedalt" style="font-size: 7pt">Injury/ill health<br/>requiring hospital<br/>admission</td>
		<td class= "hed" style="font-size: 7pt">Fatality or permanent<br/>disabling injury</td>
	</tr>
	<tr>
		<td></td>
		<td class= "hedalt" style="font-size: 7pt"><strong>Environmental<br/>consequence</strong></td>
		<td class= "hed" style="font-size: 7pt">Minor effects on<br/>biological or physical<br/>environment</td>
		<td class= "hedalt" style="font-size: 7pt">Moderate short term<br/>effects but not effecting<br/>ecosystem functions</td>
		<td class= "hed" style="font-size: 7pt">Serious<br/>medium-term<br/>environmental effects</td>
		<td class= "hedalt" colspan="2" style="font-size: 7pt">Very serious long term impairment of ecosystem<br/>functions<td>
		
	</tr>
	<TR>
		<TD class="hed">Almost Certain<br/><input title="Assess this as almost certain" type="radio" name="radiol" value="Almost Certain"  <%if strLikelyhood = "Almost Certain" then%>checked<%end if%> /></TD>
		<td class ="hed" style="font-size: 7pt;">The event will occur on<br/>an annual basis</td>
		<TD class="medium">Moderate</TD>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Critical</TD>
		<TD class="extreme">Critical</TD>
	</TR>
	<TR>
		<TD class="hedalt ">Likely<br/><input title="Assess this as likely" type="radio" name="radiol" value="Likely" <%if strLikelyhood = "Likely" then%>checked<%end if%> /></TD>
		<td class= "hedalt" style="font-size: 7pt;">The event has occurred<br/>several times or more in<br/>your career</td>
		<TD class="medium">Moderate</TD>
		<TD class="medium">Moderate</TD>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Critical</TD>
	</TR>
	<TR>
		<TD class="hed">Possible<br/><input title="Assess this as possible" type="radio" name="radiol" value="Possible"  <%if strLikelyhood = "Possible" then%>checked<%end if%> /></TD>
		<td class="hed" style="font-size: 7pt;">The event might occur<br/>once in your career</td>
		<TD class="low">Low</TD>
		<TD class="medium">Moderate</TD>
		<TD class="medium">Moderate</TD>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
	</TR>
	<TR>
		<TD class="hedalt">Unlikely<br/><input title="Assess this as unlikely" type="radio" name="radiol" value="Unlikely"  <%if strLikelyhood = "Unlikely" then%>checked<%end if%> /></TD>
		<td class="hedalt" style="font-size: 7pt;">The event does occur<br/>somewhere from time<br/>to time</td>
		<TD class="low">Low</TD>
		<TD class="low">Low</TD>
		<TD class="medium">Moderate</TD>
		<TD class="medium">Moderate</TD>
		<TD class="high">High</TD>
	</TR>
	<TR>
		<TD class="hed">Rare<br/><input title="Assess this as rare" type="radio" name="radiol" value="Rare"  <%if strLikelyhood = "Rare" then%>checked<%end if%> /></TD>
		<td class="hed" style="font-size: 7pt;">Heard of something<br/>like this occurring somewhere</td>
		<TD class="low">Low</TD>
		<TD class="low">Low</TD>
		<TD class="low">Low</TD>
		<TD class="medium">Moderate</TD>
		<TD class="medium">Moderate</TD>
	</TR>
	</TABLE>
</TD>
</TR>
</TABLE>
<p  align="left">If risk level is high or extreme, then add more control measures in step (3).</p>

  <BR />
      <hr style = "width: 82%;" align="left" />
  <div style="clear: all; "></div>

<table class="loginlist" style="margin: 0 auto">
      <tr>     
      <td><input type="submit" value="Save Risk Assessment" name="btnSave" />&nbsp;&nbsp;&nbsp;</td>
      <td><input type="submit" value="Edit SWMS" onclick="Form1.action='SWMS.asp'; return true;">&nbsp;&nbsp;&nbsp;</td>
      
      <td>
      </form>     

     <form style="display:inline;" method="post" action="DeleteQORA.asp" onSubmit="return ConfirmDelete();" />
      	<input type="submit" value="Delete Risk Assessment" name="btnDelete" />
      	<input type="hidden" name="hdnQORAId" value="<%=testval%>" />
		<input type="hidden" name="hdnFacilityId" value="<%=numFacilityID%>" />
		<input type="hidden" name="operationId" value="<%=numOperationID%>" />
		</td>
	</tr>
 </table>

      <br/>

</form>
</div>
</div>
</body>
</html>
