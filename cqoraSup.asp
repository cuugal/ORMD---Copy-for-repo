<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<%strLoginId = session("strLoginId")%>
<html>
<head>

<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <!--#include file="bootstrap.inc"--> 
<!--[if IE]>

<link rel="stylesheet" type="text/css" href="IE7.css" />

<![endif]-->
<title>Online Risk Register - Create a new Risk Assessment</title>

<SCRIPT type="text/javascript" language="Javascript" SRC="validation.js">
</SCRIPT>
<!-- Code for the hover menus -->

<SCRIPT type="text/javascript" language="Javascript" SRC="tabbed.js">
</SCRIPT>

</head>
<body>
    <!--#include file="HeaderMenu.asp" -->

<div id="wrapperform">
  <div id="content">
    <!-- new code starts here-->
    <%
Dim connAdmin
Dim rsFillAdmin
Dim connFaculty
Dim rsFillFaculty
Dim numFacultyId

'Database Connectivity Code 
  set connAdmin = Server.CreateObject("ADODB.Connection")
  connAdmin.open constr
  
   Dim strFacultyName   
   Dim strSupervisorName
   Dim strGivenName
   Dim strSurname
   
   searchType = Request.form("searchType")
   Session("searchType") = searchType
  
   '*************************** Code to get the details of location************************
    dim numFacilityId
    dim numBuildingId
    dim numCampusID
      
    numFacilityId = request.form("cboRoom")
    numBuildingId = request.form("hdnBuildingId")
    numCampusId = request.form("hdnCampusId")
    session("LastRACreatednumFacilityID") = numFacilityId 
	session("LastRACreatednumBuildingID") = numBuildingId
    session("LastRACreatednumCampusID") = numCampusId

    fromSearchQORA=request.form("fromSearchQORA")
    if(fromSearchQORA = "true" and searchType = "location") then
          set connPre = Server.CreateObject("ADODB.Connection")
          connPre.open constr
          strSQL ="Select * from tblFacility, tblBuilding  where tblFacility.numBuildingId = tblBuilding.numBuildingId and numFacilityId ="& numFacilityId
          set rsFillPre = Server.CreateObject("ADODB.Recordset")
          rsFillPre.Open strSQL, connPre, 3, 3
          numBuildingId = rsFillPre("tblBuilding.numBuildingId")
          numCampusId = rsFillPre("numCampusId")
    end if

if(searchType = "location") then   
 'code for Campus Name
  set connCampus = Server.CreateObject("ADODB.Connection")
  connCampus.open constr
' setting up the recordset
  strSQL ="Select * from tblCampus,tblBuilding  where tblCampus.numCampusId = tblBuilding.numCampusId and numBuildingId ="& numBuildingId
  set rsFillCampus = Server.CreateObject("ADODB.Recordset")
  

  rsFillCampus.Open strSQL, connCampus, 3, 3

  'Response.write(strSQL)
  strCampusName = rsFillCampus("strCampusName")

  'code for building Name
  set connBuilding = Server.CreateObject("ADODB.Connection")
  connBuilding.open constr
' setting up the recordset
   strSQL ="Select * from tblBuilding where numBuildingId ="& numBuildingId
  set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
  rsFillBuilding.Open strSQL, connBuilding, 3, 3

    strBuildingName = rsFillBuilding("strBuildingName")
  
   'code for Facility Name
  set connFacility = Server.CreateObject("ADODB.Connection")
  connFacility.open constr
' setting up the recordset
'AA jan 2010 includes fix for relationship
   strSQL ="Select * from tblFacility, tblFacilitySupervisor"_
   			&" where tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and numFacilityId ="& numFacilityId
  set rsFillFacility = Server.CreateObject("ADODB.Recordset")
   rsFillFacility.Open strSQL, connFacility, 3, 3
  
    strRoomName = rsFillFacility("strRoomName")
    strRoomNo = rsFillFacility("strRoomNumber")
    'AA jan 2010 fix for relationship
    'strLogin = rsfillfacility("strfacilitySupervisor")
    strLogin = rsfillfacility("strLoginID")
    
   strFacilityName = cstr(strRoomNo) + " - " + cstr(strRoomName)
   operationId = 0
end if
  
  'if our QORA type is operation
if(searchType = "operation") then
	set connOperation = Server.CreateObject("ADODB.Connection")
  	connOperation.open constr
	dim operationId 
	operationId = request.form("cboOperation")
	
	strSQL = "select * from tblOperations where numOperationId ="&operationId
	set rsFillOperation = Server.CreateObject("ADODB.Recordset")
	rsFillOperation.Open strSQL, connOperation, 3, 3

	dim facilitySupervisorID
	facilitySupervisorID = rsFillOperation("numFacilitySupervisorId")
	dim strOperationName 
	strOperationName= rsFillOperation("strOperationName")
	
	strSql = "select * from tblFacilitySupervisor where numSupervisorID = "&facilitySupervisorID
	set rsFillFacilitySuper = Server.CreateObject("ADODB.Recordset")
	rsFillFacilitySuper.Open strSQL, connOperation, 3, 3
	
	strGivenName = rsFillFacilitySuper("strGivenName")
	strSurname = rsFillFacilitySuper("strSurname")
	strSupervisorName = cstr(strGivenName) +" "+ cstr(strSurname)  
	strLogin = rsFillFacilitySuper("strLoginID")
	
	numFacultyID = rsFillFacilitySuper("numFacultyID")
	
	strSQL = "select strFacultyName from tblFaculty where numFacultyID = "&numFacultyID
	set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
	rsFillFaculty.Open strSQL, connOperation, 3, 3
	
	strFacultyName = rsFillFaculty("strFacultyName")   
 	numFacilityID =0
end if


strSQL ="Select * from tblFacilitysupervisor where strLoginId = '"& strLogin &"'"


set rsFillAdmin = Server.CreateObject("ADODB.Recordset")
rsFillAdmin.Open strSQL, connAdmin, 3, 3

if rsFillAdmin.EOF = False then
  strGivenName = rsFillAdmin("strGivenName") 
  strSurName = rsFillAdmin("strSurName") 
  
  strSupervisorName = cstr(strGivenName) +" "+ cstr(strSurname)  
  numFacultyId = rsFillAdmin("numFacultyId")

else
response.write("exception caught !")  
end if

'Database Connectivity Code 
  set connFaculty = Server.CreateObject("ADODB.Connection")
  connFaculty.open constr
  
' setting up the recordset

strSQL ="Select * from tblFaculty where numFacultyId ="&numFacultyId

set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
rsFillFaculty.Open strSQL, connFaculty, 3, 3

strFacultyName = rsFillFaculty("strFacultyName")

'separate connection to db to obtain the proposed new RA number
set dcnDb2 = server.CreateObject("ADODB.Connection")
dcnDb2.Open constr
set rsSearch2 = server.CreateObject("ADODB.Recordset")
strSQL2 = "Select max(numQORAId)+1 as newQORAId from tblQORA" 
rsSearch2.Open strSQL2, dcnDb2, 3, 3
strNewQORAId = rsSearch2("newQORAId")
strJobSteps = ""
%>
    
    <form method="post" 	action="AddCQORAsup.asp" name="Form1" onSubmit="return ConfirmChoice();">
      <input type="hidden" name="hdnBuildingId"  	value="<%=numBuildingID%>" />
      <input type="hidden" name="hdnCampusId" 		value ="<%=numCampusID%>"/ >
      <input type="hidden" name="hdnFacilityId" 	value ="<%=numFacilityID%>" />
      <input type="hidden" name="hdnFacultyId" 		value ="<%=numFacultyID%>" />
      <input type="hidden" name="hdnLoginId" 		value ="<%=strLogin%>" />
      <input type="hidden" name="hdnFacilityName" 	value ="<%=strFacilityName%>" />
      <input type="hidden" name="operationId" 		value ="<%=operationId%>" />
      <input type="hidden" name="searchType" 			value="<%=searchType%>" />
      <input type="hidden" name="hasSWMS" 			value="" />
      <table width = 80%>
     	<tr>
      		<td align="left"><h2 class="pagetitle">Enter the details of the new Risk Assessment</h2></td>
      		<td align="right"> <h2> RA Number <%=strNewQORAId%></h1></td>
      	</tr>
      </table>
      
      <table class="suprreportheader" style="width: 82%;">
        <input type="hidden" name="VTI-GROUP" value="1" />
        <tr>
          <th>Faculty/Unit</th>
          <td colspan="3"><%=strFacultyName%></td>
        </tr>
        <tr>
        
         <%
        if(searchType = "location") then %>
        
        <th>Facility</th>
          <td><strong>Campus</strong><br/><%=strCampusName%></td>
          <td><strong>Building</strong><br/><%=strBuildingName%></td>
          <td><strong>Room Number/Name</strong><br/><%=strFacilityName%></td>
        </tr>
        <% end if
        if(searchType = "operation") then %>
	        <th>Operation</th>
	          <td colspan="3"><%=strOperationName%></td>
	        </tr>
        <% end if %>
        <tr>
          <th>Supervisor Name</th>
          <td colspan="3"><%=strSupervisorName%></td>
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
        	<th>Assessor / Reviewer:</th><td colspan="3"><input type="text" name="txtAssessor" size="35" /></td>
        	<!--<td colspan="2">Date Last Modified (dd/mm/yyyy)&nbsp;&nbsp;&nbsp;<input type="text" name="txtDateCreated" size="9" value="<%=todaysDate%>"/></td>-->
            <input type="hidden" name="txtDateCreated" size="9" value="<%=todaysDate%>"/>
			<!-- the line above is needed as an input box to prevent error - how can this be converted to plain text as attempted in line below-->
        	<!--td colspan="2">Date Last Modified (dd/mm/yyyy)&nbsp;&nbsp;&nbsp;<%Response.write(todaysDate)%></td-->
          	<!--td>Review Date&nbsp;&nbsp;&nbsp;<%=renewalDate%></td-->
        </tr>
		<tr>
			<th>Persons Consulted</th>
			<td colspan="3"><textarea rows="1" name="strConsultation" cols="90" ></textarea></td>
		</tr>
      </table>
    </br>
	  <table class="reportwarn" style="width: 82%">
      <tr><td><b>NOTE: All risk assessments should be performed in consultation with staff involved with this work activity.</b></td></tr>
	  </table>
	</br>

      	<strong>(1) Describe briefly how this work is performed</strong><br>
      <br>
      <table class="suprreportheader" style="width: 82%">
        <tr>
          <th>Work Activity Description:</th>
          <td><!--<input type="text" name="txtTaskDesc" size="100%" />-->
          <textarea rows="4" name="txtTaskDesc" cols="90" ></textarea></td>
        </tr>
      </table>
      <!-- <hr /> -->
      <!-- relocated -->
      <br>
      <hr style = "width: 82%;" align="left" />
      
      
     <!--Navigation -->
  <a name="point1"></a> <strong>(2) Select hazards relating to the work activity.</strong>
  <p style="width: 82%">Select from the menu below all of the hazards that apply to the work.<br />
  NOTE: Lists of hazards appear when you put the cursor over this menu. When you click on one it appears in the text box below.</p>
<div>
 <table class="suprreportheader" style="width: 82%">
 	<tr>
		<td colspan="2">


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
               <li><a href="#point1" onClick="Populate('Chemical - Hazardous Substances or Dangerous Goods\r\n'); " title="Click to add 'Hazardous Substances or Dangerous Goods' as a Hazard in this Risk Assessment.">Hazardous Substances or Dangerous Goods</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Hazardous Waste\r\n'); " title="Click to add 'Hazardous Waste' as a Hazard in this Risk Assessment.">Hazardous Waste</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Fumes\r\n');" title="Click to add 'Fumes' as a Hazard in this Risk Assessment.">Fumes</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Dust\r\n'); " title="Click to add 'Dust' as a Hazard in this Risk Assessment.">Dust</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Chemical - Vapours\r\n'); " title="Click to add 'Vapours' as a Hazard in this Risk Assessment.">Vapours</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Gases\r\n'); " title="Click to add 'Gases' as a Hazard in this Risk Assessment.">Gases</a></li>
               <li><a href="#point1" onClick="Populate('Chemical - Fire/Explosion Risk\r\n'); " title="Click to add 'Fire/Explosion Risk' as a Hazard in this Risk Assessment.">Fire/Explosion Risk</a></li>
            </ul>
         </div>
      </div>
   </div>
   <div id="biological" class="tab-pane fade">
      <div class="row">
         <div class="col-xs-4">
            <ul class="section-list">
               <li><a href="#point1" onClick="Populate('Biological - Imported Biomaterials\r\n')" title="Click to add 'Imported Biomaterials' as a Hazard in this Risk Assessment.">Imported Biomaterials</a></li>
               <li><a href="#point1" onClick="Populate('Biological - Cytotoxins\r\n')" title="Click to add 'Cytotoxins' as a Hazard in this Risk Assessment.">Cytotoxins</a></li>
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

	<!--tr>
    	<td>List hazards below either by using the menu above or typing directly into the text box.</td>
      	<td>Describe here inherent risks of the work: <br/>
      	<strong><u>How</u></strong> these hazards cause harm<br/>
      	<strong><u>What</u></strong> sort of injury/illness might occur?</td>

	</tr-->

	<tr>
    	<td><strong>Hazard List</strong> <br/> List hazards below either by using the  menu above</br> or typing directly into the text box.</td>
      	<td><strong>Inherent Risks</strong> <br/>List injury/illness that could occur from this work activity and 
      	<strong><u>how</u></strong> this harm could happen. </br><i>e.g. - Back strain may occur from incorrect lifting.</i></td>
	</tr>



  	<tr>  
    	<td><!-- textarea box goes in this table cell -->
        <br/>
        	<textarea rows="8" name="T1" id="T1" cols="45"></textarea>
            
        </td>
        <td><!-- textarea box goes in this table cell -->
        	<br />
          	<textarea rows="8" name="T3" id="T3" cols="45" ></textarea>
		</td>
		</tr>

          
       
		<!-- </table>  -->
	</table>
	</br>
	<table class="reportwarn" style="width: 82%">
		<tr>
			<td colspan="2"> Do the hazards you have selected have the potential to cause death, or serious injury / illness (causing temporary disability)? &nbsp 
				<strong>YES</strong> 
				<input type="radio" name="boolSWMSRequired" Value="Yes" />  &nbsp &nbsp 
				<strong>NO</strong> 
				<input type="radio" name="boolSWMSRequired" Value="No"  /><br />
				If they do, then a Safe Work Method Statement (SWMS) must also be recorded.	
			</td>
		</tr>
	</table>

 </div>
    <div style="clear:both">
    </div> 
 
     
      <br>
      <hr style = "width: 82%;" align="left" />
      
      
      <strong>(3) Select safety control measures to make this work safe.</strong>
      <p style = "width: 82%;">- Select the safety control measures needed to minimise the risk of harm to an acceptable level. Refer to <a href="http://www.fsu.uts.edu.au/procurement/staff-only/form.html"  target="_blank">FSU purchasing policy and procedures</a> where cost considerations may impact on control selection.</br>
	  - List the Safety Control Measures that are both 'currently in place' and 'proposed'.</br>
	  NOTE: Lists of safety control measures appear when you put the cursor over this menu.</p>
      
<div>
	<!--********old menu archived - contact Andrew Alger -->
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
               <li><a href="#point2" onClick="PopulateNext('- Guarding/Barriers\r\n')" title="Add the control 'Guarding/Barriers' to this Risk Assessment.">Guarding / Barriers</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Biosafety Cabinet\r\n')" title="Add the control 'Biosafety Cabinet' to this Risk Assessment.">BioSafety Cabinet</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Fume Cupboard/Local Exhaust Ventillation\r\n')" title="Add the control 'Fume Cupboard/Local Exhaust Ventillation' to this Risk Assessment.">Fume Cupboard / Local Exhaust Ventilation</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Redesign the Workspace/Workflow\r\n')" title="Add the control 'Redesign the Workspace/Workflow' to this Risk Assessment.">Redesign the Workspace / Workflow</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Lifting Equipment/Trolleys\r\n')" title="Add the control 'Lifting Equipment/Trolleys' to this Risk Assessment.">Lifting Equipment / Trolleys</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Regular Maintenance of Equipment\r\n')" title="Add the control 'Regular Maintenance of Equipment' to this Risk Assessment.">Regular Maintenance of Equipment</a></li>
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
               <li><a href="#point2" onClick="PopulateNext('- Chemical Risk Assessment\r\n')" title="Click to add the administrative control 'Chemical Risk Assessment' to this Risk Assessment.">Chemical Risk Assessment</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Licensing/Certification of Operators\r\n')" title="Click to add the administrative control 'Licensing/Certification of Operators' to this Risk Assessment.">Licensing/Certification of Operators</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Test and Tag Electrical Equipment\r\n')" title="Click to add the administrative control 'Test and Tag Electrical Equipment' to this Risk Assessment.">Test and Tag Electrical Equipment</a></li>
            </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Monitor Exposure Level (Sound/Substance/Radiation)\r\n')" title="Click to add the administrative control 'Monitor Exposure Level (Sound/Substance/Radiation)' to this Risk Assessment.">Monitor Exposure Level</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Licences (Lifts, Boilers, Pressure Vessles, Radiation)\r\n')" title="Click to add the administrative control 'Licences (Lifts, Boilers, Pressure Vessles, Radiation)' to this Risk Assessment.">Licences (Lifts, Boilers, Pressure Vessels, Radiation)</a></li>
               <li><a href="#point2" onClick="PopulateNext('- Biosafety Committe Assessment (GMOs, pathogens, radiation, cytotoxins, imported biologicals)\r\n')" title="Click to add the administrative control 'Biosafety Committe Assessment (GMOs, pathogens, radiation, cytotoxins, imported biologicals)' to this Risk Assessment.">BioSafety Committee Assessment (GMOs, pathogens, radiation, cytotoxins, imported biologicals)</a></li>
               <li><a href="#point2" onClick="PopulateNext('- UTS Fieldwork Guidelines for overnight excursions in the field\r\n')" title="Click to add the administrative control 'UTS Fieldwork Guidelines for overnight excursions in the field' to this Risk Assessment.">UTS Fieldwork Guidelines for overnight excursions in the field</a></li>
             </ul>
         </div>
         <div class="col-xs-4">
            <ul>
               <li><a href="#point2" onClick="PopulateNext('- Regular Breaks & Task Rotation\r\n')" title="Click to add the administrative control 'Regular Breaks & Task Rotation' to this Risk Assessment.">Regular Breaks & Task Rotation</a></li>
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
               <li><a href="#point2" onClick="PopulateNext('- Gloves\r\n')" title="Click to add the risk control 'Gloves' to this Risk Assessment.">Gloves</a></li>
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
               <li><a href="#point2" onClick="PopulateNext('- Protective Clothing/Apron/Overalls\r\n')" title="Click to add the risk control 'Protective Clothing/Apron/Overalls' to this Risk Assessment.">Protective Clothing / Apron / Overalls</a></li>
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
		<th colspan="2" style="text-align:center"><strong>Currently In Place &nbsp &nbsp OR &nbsp &nbsp  Proposed Implementation Date</strong></th>
		<th style="text-align:center"><strong>Remove</strong></th>
	</tr>
    <tr>
		<td colspan="1" style="font-size: 7pt;text-align:center">If you cannot find the desired control in the menu,<br/>you can add your own by clicking 'Add Row'.</td>
       	<td style="font-size: 7pt;text-align:center">Tick the checkbox when the safety control measure is in place.</td>
		<td style="font-size: 7pt;text-align:center">You must enter proposed implementation date for proposed safety control measures.</td>
       	<td style="font-size: 7pt;text-align:center">Click remove to delete this row.</td></tr>       		
	</tr>
        	
         	  	
		</table>
       	<table class="bluebox" style="width:82%;">

       		<tfoot><tr><td style="text-align:right"></td>
       			<!--<input type="button" value="Remove" onclick="removeRowFromTable();" />-->
       			
       			<td style="min-width:240px" colspan="1">&nbsp;&nbsp;&nbsp;</td>
       			
       			<td style="text-align:right">
       			<input type="button" value="Add Row" onclick="addRowToTable();" style="margin-right:85px" />
       			<input type="button" value="Remove" onclick="deleteRow();" style="margin-right:35px" /></td>
       		</tr></tfoot>
       		</table>
      <div style="clear:both">
      </div> 
   </div>
  
       
      <hr style = "width: 82%;" align="left" />


<strong>(4) Assess level of residual risk</strong>
  <p style="width: 82%">- Use the risk matrix below as a guide to assess the level of risk, based on the hazards identified above and the way that the work is done with safety control measures that are in place. </br>
- High or Extreme risk is not acceptable. To reduce likelihood / consequence, add more control measures in step (3).</p>

<!-- Risk Matrix -->


<!--TABLE style="margin-left:160px"-->
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
		<TD class="hed"><input title="Assess this as insignificant" type="radio" name="radioc" value="Insignificant" onClick="return radiounClick (event)"  checked/>Insignificant</TD>
		<TD class="hedalt"><input title="Assess this as minor" type="radio" name="radioc" value="Minor" onClick="return radiounClick (event)"  />Minor</TD>
		<TD class="hed"><input title="Assess this as moderate" type="radio" name="radioc" value="Moderate" onClick="return radiounClick (event)"  />Moderate</TD>
		<TD class="hedalt"><input title="Assess this as major" type="radio" name="radioc" value="Major" onClick="return radiounClick (event)"  />Major</TD>
		<TD class="hed"><input title="Assess this as catastrophic" type="radio" name="radioc" value="Catastrophic" onClick="return radiounClick (event)"  />Catastrophic</TD>
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
		<TD class="hed">Almost Certain<br/><input title="Assess this as almost certain" type="radio" name="radiol" value="Almost Certain" onClick="return radiounClick (event)"  checked/></TD>
		<td class ="hed" style="font-size: 7pt;">The event will occur on<br/>an annual basis</td>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Extreme</TD>
		<TD class="extreme">Extreme</TD>
		<TD class="extreme">Extreme</TD>
	</TR>
	<TR>
		<TD class="hedalt ">Likely<br/><input title="Assess this as likely" type="radio" name="radiol" value="Likely" onClick="return radiounClick (event)"  /></TD>
		<td class= "hedalt" style="font-size: 7pt;">The event has occurred<br/>several times or more in<br/>your career</td>
		<TD class="medium">Medium</TD>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Extreme</TD>
		<TD class="extreme">Extreme</TD>
	</TR>
	<TR>
		<TD class="hed">Possible<br/><input title="Assess this as possible" type="radio" name="radiol" value="Possible" onClick="return radiounClick (event)"  /></TD>
		<td class="hed" style="font-size: 7pt;">The event might occur<br/>once in your career</td>
		<TD class="low">Low</TD>
		<TD class="medium">Medium</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Extreme</TD>
		<TD class="extreme">Extreme</TD>
	</TR>
	<TR>
		<TD class="hedalt">Unlikely<br/><input title="Assess this as unlikely" type="radio" name="radiol" value="Unlikely" onClick="return radiounClick (event)"  /></TD>
		<td class="hedalt" style="font-size: 7pt;">The event does occur<br/>somewhere from time<br/>to time</td>
		<TD class="low">Low</TD>
		<TD class="low">Low</TD>
		<TD class="medium">Medium</TD>
		<TD class="high">High</TD>
		<TD class="extreme">Extreme</TD>
	</TR>
	<TR>
		<TD class="hed">Rare<br/><input title="Assess this as rare" type="radio" name="radiol" value="Rare" onClick="return radiounClick (event)"  /></TD>
		<td class="hed" style="font-size: 7pt;">Heard of something<br/>like this occurring somewhere</td>
		<TD class="low">Low</TD>
		<TD class="low">Low</TD>
		<TD class="medium">Medium</TD>
		<TD class="high">High</TD>
		<TD class="high">High</TD>
	</TR>
	</TABLE>
</TD>
</TR>
</TABLE>

<!--
<TABLE>
	<TR>
	<TD>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</TD>
	<TD><!-- original three button risk selector 
	  <table class="eq">
		<tr>
		<td class="low">Low<input title="Assess this task as low risk." type="radio" name="radios" value="Third" onClick="return radiounClick (event)" checked /></td>
		<td class="medium">Medium<input title="Assess this task as medium risk." type="radio" name="radios" value="Second" onClick="return radiounClick (event)" /></td>
		<td class="high">High<input title="Assess this task as high risk." type="radio" name="radios" value="First" onClick="return radioClick (event)" /></td>
		<td><FONT SIZE="1">Extreme residual risk means <br>insufficient control measures<br> in place.<br></FONT></TD>
		</TR>
	  </table>
	</TD>
	</TR>
</TABLE>-->

      <BR />
      <hr style = "width: 82%;" align="left" />
      
      <div style="clear: all; "></div>
      
      
 
      <!--
Come back and fill in this date when actions have been completed.<br />
      <input type="text" name="txtDtActionsCompleted" size="10" />
      <span style="font-size: 8pt;">(Format: dd/mm/yyyy)</span> <br />
 -->
      <div class="loginlist">
      <ul>
      	<!--SWMS doesn;t exist until RA has been completed
      	<li><input type="submit" value="SWMS" name="btnSWMS" />&nbsp;&nbsp;&nbsp;</li>-->
      	<li><input type="submit" value="Save Risk Assessment" name="btnSave" />&nbsp;&nbsp;&nbsp;</li>
      	<li><input type="submit" value="Create SWMS" onclick="Form1.action='SWMS.asp'; return true;"></li>

      	</ul>
      </div>
      
    </form>

      
    <br />
    </center>
  </div>
</div>
</body>
</html>