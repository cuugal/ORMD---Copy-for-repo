<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!--#include file="aspJSON.asp" -->
<%strLoginId = session("strLoginId")%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<link rel="SHORTCUT ICON" href="favicon.ico" type="image/x-icon" />

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

	strPPE = rsResults("ppe")
	strEq = rsResults("emergency")
	
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
	

    dim canEdit
    canEdit = false
    'If we aren't logged in then you can't edit
    if session("LoggedIn") then
        ' if we are admin then we can edit
        if session("isAdmin") then
            canEdit = true
        else
            ' otherwise, check the DB to see if we have ownership of this record, ergo edit permission
            set rsCanEdit = server.CreateObject("ADODB.Recordset")
            strSQL = " Select count(*) as editable from ( "_
                &" Select tblFacility.numFacilitySupervisorId from tblQORA, tblFacility"_
                &" where tblQORA.numFacilityId = tblFacility.numFacilityId"_
                &" and numQORAID = "&testval&" and numFacilitySupervisorId = "&session("numSupervisorId")_

               &" union all "_
                &" Select tblOperations.numFacilitySupervisorId from tblQORA, tblOperations"_
                &" where tblQORA.numOperationId = tblOperations.numOperationId"_
                &" and numQORAID = "&testval&"  and numFacilitySupervisorId = "&session("numSupervisorId")_
                &")"
           
            rsCanEdit.Open strSQL, dcnDb, 3, 3
            if cint(rsCanEdit("editable")) > 0 then
                canEdit = true
            end if
        end if
    end if
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />

     <!--#include file="bootstrap.inc"--> 
</head>
<body>
	<!--#include file="HeaderMenu.asp" -->
<title>EHS Risk Register for Facilities - Edit a Risk Assessment</title>

<div id="wrapperform">
  <div id="content">	
  <table width = 80%>
     	<tr>
      		<td align="left"><h2 class="pagetitle">Safe Work Method Statement (SWMS)</h2> 
   
  </td>
      		<td align="right"> <h2> Risk Assessment Number <%=testval%></h2></td>
      	</tr>
      </table>
      
   <!--<form target="_blank" method="post" action="SWMS-print.asp"> -->
     <form method="post" name="Form1" action="AddSWMS.asp">  
       
 <table class="swms" width = 80%>
     	<tr>
      		<td align="left" class="plainbox" colspan="4"><strong>Work Activity Description:</strong> <%=rsResults("strTaskDescription")%></td>
		</tr>
		<tr>
			
			<% 'only display save button if the user is logged in, and has write access to the record, or is a admin.
				If canEdit Then %>
					
					<td> 
					<input type="submit" value="To Risk Assessment" onclick="Form1.action='EditQORA.asp?numCQORAId=<%=testval%>'; Form1.target='_self'; return true;">
					</td>

					<td >
					<!-- 9April2010 DLJ added target=self -->
					<input type="submit" value="Save SWMS" target="_self" onclick="Form1.action='AddSWMS.asp'; Form1.target='_self'; return true;"/>
					<!--</form>-->
					</td>
					
			    <% End If %>
					
                  <%
                      dim action
                      if Session("mostRecentSearch") <> "" then
                      %>
                         <td>
					        <input type="button" value="Back to Risk Assessment List" name="Back" onclick="$('#refreshResults').submit();">
                        </td>
                        <%
                        action = Session("mostRecentSearch")
                      else
                        action = "Home.asp"
                      end if
                       %>

                     


					

      				<td align="center">
					<!-- <input type="submit" value="Print preview" /> -->		
					<input type="submit" value="Print Preview" target="_blank" onclick="Form1.action='SWMS-print.asp'; Form1.target='_blank'; return true;" />			
					<input type="hidden" name="hdnQORAId" value="<%=testval%>" /> 
					<input type="hidden" name="hdnFacilityId" value="<%=numFacilityID%>" />
					<input type="hidden" name="operationId" value="<%=numOperationID%>" />
					<input type="hidden" name="canEdit" value="<%=canEdit%>"/>
	  				</td>
		</tr>
  </table>
 <table class="swms" width = 80%>
	<tr>
      		<% if strAssessRisk="L" then%> <td class="low" align="middle" width="250"> <strong>Residual Risk Level: Low </strong></td><%end if%>
      		<% if strAssessRisk="M" then%> <td class="medium" align="middle" width="250"> <strong>Residual Risk Level: Medium</strong> </td><%end if%>
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


      <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
		<tr>  
			<td>
				<strong>HAZARDS:</strong><br> <%=strHazardsDesc%>          
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
  
  <table class="bluebox" style="margin: 0; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
           <!--#include file="pictogram.asp" -->
            <!--textarea rows = "8" style="width:100%;" name="T3" readonly-->
<% 'here we need to populate the textarea with any existing controls we can locate

            Set ppeImageURLs = CreateObject("System.Collections.ArrayList")
            Set eqImageUrls = CreateObject("System.Collections.ArrayList")

        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""

        	while not rsControls.EOF 

         		'get the images we need to display
         		dim thisControl
         		thisControl = rsControls("strControlMeasures")
                thisControl = Replace(thisControl, "-", "")
                thisControl = Trim(thisControl)

                strControlsImplemented = strControlsImplemented &rsControls("strControlMeasures")& "<BR>"

     		' get the next record
           rsControls.MoveNext
     		wend


     		%>
     		<%=strControlsImplemented%>
<!--/textarea-->           
</td>
</tr> 
</table>

<script type="text/javascript">
    var ppeItems = [];
    var eqItems = [];

    function addPPE(){
        ppeItems = [];
        $('input.ppeClass:checked').each(function(){
            ppeItems.push($(this).attr('name'));
        });
        document.getElementById("ppe").value = JSON.stringify(ppeItems);
    }
    function addEq(){
         eqItems = [];
         $('input.eqClass:checked').each(function(){
            eqItems.push($(this).attr('name'));
         });
        document.getElementById("eq").value = JSON.stringify(eqItems);
    }

</script>


<strong>PPE Required for this activity</strong>
<table class="bluebox" style="margin: 0; width:80%; padding-left:40px">
  <tr><td>
  <%
  For Each key In ppe.keys
  %>
  <div style="float:left;padding-right:5px" align="center">
        <image width="100px" src="images/<%=ppe.item(key)%>"/><br/>
        <input type="checkbox" class="ppeClass" name="<%=key%>" onclick="addPPE();"/>
  </div>
  <%
  Next
  %>
</td></tr>

</table>
<br/>       
<strong>Emergency Equipment required for this activity</strong>
<table class="bluebox" style="margin: 0; width:80%; padding-left:40px">
 <tr><td>
   <%
   For Each key In eq.keys
   %>
   <div style="float:left;padding-right:5px" align="center">
        <image width="100px" src="images/<%=eq.item(key)%>"/><br/>
        <input type="checkbox" class="eqClass" name="<%=key%>" onclick="addEq();"/>
   </div>
   <%
   Next
   %>
 </td></tr>
</table>
                            <input type="hidden" name="ppe" id="ppe" value=""/>
                          <input type="hidden" name="eq" id="eq" value=""/>
<script type="text/javascript">

    str1 = '<%=strppe%>';
    str2 = '<%=streq%>';

    if(str1 != ''){
         ppeItems =  JSON.parse(str1);
    }

    if(str2 != ''){
         eqItems =  JSON.parse(str2);
    }


    $('input.ppeClass').each(function(){

        if($.inArray($(this).attr('name'), ppeItems)!== -1){
            $(this).prop( "checked", true );
        }
        else{
            $(this).prop( "checked", false );
        }
     });
     $('input.eqClass').each(function(){
             if($.inArray($(this).attr('name'), eqItems)!== -1){
                 $(this).prop( "checked", true );
             }
             else{
                 $(this).prop( "checked", false );
             }
      });
      $('#ppe').val(JSON.stringify(ppeItems));
      $('#eq').val(JSON.stringify(eqItems));
</script>


<br/>
<% 

if isNull(strJobSteps) Then
	strIntro = "The default wording below is a guide only. Please edit and delete example text as needed."
	strJobSteps = "  BEFORE YOU START:"&vbCrLf&" - Training or qualification required"&vbCrLf&" - Safety equipment and PPE required"&vbCrLf&" - Allocated responsibilities"_
	&vbCrLf&vbCrLf&"  STEP BY STEP PROCEDURE:"&vbCrLf&" - Pre-check and inspect equipment"&vbCrLf&" - Setting up the activity - e.g. Measure out materials"_

	&vbCrLf&vbCrLf&"   Mechanical equipment example steps e.g. setting up job, starting machine"_

	&vbCrLf&"   Microbiological example steps e.g. culturing liquid, streaking plates, centrifugation, long-term storage, transporting"_

	&vbCrLf&vbCrLf&"  FINISHING UP:"&vbCrLf&" - Clean up / decontamination"_

	&vbCrLf&vbCrLf&"  WASTE DISPOSAL:"&vbCrLf&" - Liquid waste disposal"&vbCrLf&" - Solid waste disposal"_

	&vbCrLf&vbCrLf&"  MAINTENANCE:"_

	&vbCrLf&vbCrLf&"  EMERGENCY PROCEDURES:"&vbCrLf&" - Dial 6 in emergency"&vbCrLf&" - Spill kit and spill response"&vbCrLf&" - Fire control"_

	&vbCrLf&vbCrLf&"  CODES AND STANDARDS APPLICABLE:"

   end if %>
 <strong> Work Activity steps </strong> <br/>&mdash; These "Work Activity Steps" can be edited directly. Refer to control measures listed above. 
 <br/><strong><%=strIntro%></strong>

  <table class="bluebox" style="margin 0 auto; width:80%; padding-left:40px">
  
      	<tr>  
          <td>
            <textarea rows = "40" style="width:100%;" id="T4" name="T4" ><%=strJobSteps%>
</textarea> 
          
</td>
 </tr> 
</table>

  </form>

  <% 'only allow edit if the user is logged in, and has write access to the record, or is a admin.
  				If not canEdit Then %>
  				<script type="text/javascript">
  				    $('input.ppeClass').each(function(){
                        $(this).prop( "disabled", true );

                     });
                     $('input.eqClass').each(function(){
                        $(this).prop( "disabled", true );

                      });
                     $('#T4').prop('readonly', true);

  				</script>
  				<% end if %>

       <form id="refreshResults" action="<%=action %>" method="post">
                          <input type="hidden" name="confirmationMsg" value="" />
                        <input type="hidden" name="searchType" value="<%=session("searchType") %>" />
                        <input type="hidden" name="cboOperation" value="<%=session("cboOperation")  %>" />
                        <input type="hidden" name="cboFacility" value="<%=session("cboFacility") %>" />
                        <input type="hidden" name="hdnFacultyId" value="<%=session("cboFaculty") %>" />
                        <input type="hidden" name="hdnFacilityId" value="<%=session("hdnFacilityId") %>" />
                          <input type="hidden" name="hdnBuildingId" value="<%=session("hdnBuildingId") %>" />
                          <input type="hidden" name="hdnCampusId" value="<%=session("hdnCampusId") %>" />
                        <input type="hidden" name="txtHazardoustask" value="<%=session("hdnHTask") %>" />
                        <input type="hidden" name="cboSupervisorName" value="<%=session("cboSupervisorName") %>" />
                    </form>
    <div style="clear:both"></div>  
  	</div>
  </div>
  

<br/>
<br/>
  </body>
  </html>