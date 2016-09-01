<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<link rel="SHORTCUT ICON" href="favicon.ico" type="image/x-icon" />
<%
dim loginId
loginId = session("strLoginId")
testval = request.form("hdnQORAID")
'Response.Write(loginId)
%>
<script src="bootstrap/js/jQuery-1.11.3.min.js"></script>
<script src="bootstrap/js/bootstrap.min.js"></script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!--<link rel="stylesheet" type="text/css" href="orr.css" media="screen,print" />-->
 <link rel="stylesheet" type="text/css" href="orrprint.css" media="screen,print" />
 <title>Online Risk Register - SWMS Report for Printing</title>
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
if sString <> "" then
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
end if
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
  numPageStatus = request.querystring("cboValDummy")
  numOptionId = Request.QueryString("numOptionID")
  canEdit = Request.form("canEdit")
  'Response.write(numOptionID)
      
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr


  '*********************writting the SQL ******************************
      
  '------------------------get the faculty for the login ---------------
  strSQL = "Select * "_
  &" from tblfacilitySupervisor,tblFaculty "_
  &" where tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId "_
  &" and tblFacilitySupervisor.strLoginId = '"& loginId &"'" 
  Response.write(SQL)
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty.Open strSQL, Conn, 3, 3     
  %>

<body>

<div id="wrapper">

<div id="content">

<!-- outside table -->
<table class="mainprintable" >
	<tr>
		<td>
			<img src="utslogo.gif" width="184" height="41" alt="" align="left" />
		</td>
		<td align="center">
			<h1>SAFE WORK METHOD<br />STATEMENT (SWMS)</h1>
		</td>
		<td align="right" class="small-font"> 
			University of Technology, Sydney<br />
			P.O. Box 123<br />
			Broadway NSW 2007<br />
			Australia<br />
			Based on risk assessment no. <%=testval%><br />
		</td>
	</tr>

     	

 <%

 Function FormatDate(input)
     FormatDate = Day(CDate(input)) &" "& MonthName(Month(CDate(input)))&" " & (Year(CDate(input))-1)
 End Function

	testval = request.form("hdnQORAID")
	'Response.write(testval)
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

	strFacultyName = rsF("strFacultyName") 
	 
%>
		<tr>
      		<td colspan="3" align="left"> <strong>Work Activity Description: </strong></td>
      	</tr>
      	<tr><td colspan="3" class="box"><%=rsResults("strTaskDescription")%></td></tr>
		<tr>
      		
      		<% if strAssessRisk="L" then%> <td colspan="3" class="low box" align="center" width="250"><strong>Residual Risk Level: LOW</strong></td><%end if%>
      		<% if strAssessRisk="M" then%> <td colspan="3" class="medium box" align="center" width="250"><strong>Residual Risk Level: MEDIUM</strong> </td><%end if%>
      		<% if strAssessRisk="H" then%> <td colspan="3" class="high box" align="center" width="250"><strong>Residual Risk Level: HIGH</strong></td><%end if%>
      		<% if strAssessRisk="E" then%> <td colspan="3" class="extreme box" align="center" width="250"><strong>Residual Risk Level: EXTREME<strong></td><%end if%>
      	</tr>

		<tr>
 			<td colspan="3">

 <table class="suprlevel-print" style="width: 100%;">
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
		%>
		<tr>    		
    		<td class="campus" colspan="2"><strong>Campus: </strong><%=rsc("strCampusName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="2"><strong>Building: </strong><%=rsb("strBuildingName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="2"><strong>Room Name: </strong><%=strRoomName%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="2"><strong>Room Number: </strong><%=strRoomNo%>&nbsp;&nbsp;&nbsp;</td>
  		</tr>
  		<tr>
  			<td class="campus" colspan="2"><strong>Supervisor: </strong><%=strSupervisor%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="6"><strong>Faculty: </strong><%=strFacultyName%>&nbsp;&nbsp;&nbsp;</td>		
  		</tr>
<% end if 
if(rsResults("numOperationId") <> 0) then %>
		<tr>
			<td class="campus" colspan="2"><strong>Supervisor: </strong><%=strSupervisor%>&nbsp;&nbsp;&nbsp;</td>	
	        <td class="campus" colspan="6"><strong>Operation:</strong><%=rsOper("strOperationName")%>&nbsp;&nbsp;&nbsp;</td>	
	    </tr>
<% end if %>

  		<tr>
  			<td class="campus" colspan="2">
  			
			<%' Code to create an Australian date format
			todaysday = day(date)
			todaysMonth = month(date)
			todaysYear = year(date)
			renewal = todaysYear + 1

			todaysDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(todaysYear)
			renewalDate = cstr(todaysDay) +"/"+cstr(todaysMonth)+"/"+cstr(renewal)
			%>
        	<Strong>Assessor:</strong> <%=strAssessor%></td>
        	<td class="campus" colspan="6"><strong>Date Last Modified (dd/mm/yyyy):</strong>&nbsp;&nbsp;&nbsp;<%=FormatDate(dtDateCreated)%></td>
        </tr>
	</table>

	<p class="small-font"> Review SWMS when change to work activity or at least annually. Changes can be recorded on www.orr.uts.edu.au</p>
	
	<table class="suprlevel-print" style="width: 100%;">
			<tr>
				<td>Review No</td><td>01</td><td>02</td><td>03</td><td>04</td><td>05</td><td>06</td><td>07</td><td>08</td><td>09</td>
			</tr>
				<td>Initial:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
			<tr>
				<td>Date:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
			</tr>
	</table>

	</td>
	</tr>

	<tr>
	<td colspan = "4">
	<br/>
	<strong>Hazards Relating to this Work Activity</strong>
	<table class="suprlevel-print" >
		<tr>
		<td style="width: 100%;">
		<strong>HAZARDS: </strong><br/>
<%=Escape(strHazardsDesc)%><br/>

		</td>
		</tr>
	</table>
	</td>
	</tr>
	<tr>
		<td colspan = "4">
    	<br/>
    	<strong>CAUTION !</strong>
    	<table class="suprlevel-print">
		<tr>
		<td class="medium box" style="width: 100%;">
		<strong>
<%=Escape(strInherentRisk)%><br/>
	</strong>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	
	<tr>
	<td colspan = "4">
	<br/>
	<strong>Control Measures (Safety Equipment, Training, Signage and Information)</strong>
	<table class="suprlevel-print">
		<tr>
		<td class="suprlevel-print" style="width: 100%;">
		<!--#include file="pictogram.asp" -->

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
<%=Escape(strControlsImplemented)%><br/>

		</td>
		</tr>
		</table>
	</td>
	</tr>

	<tr>
        <td colspan = "4">
        <br/>
        <strong>PPE Required for this activity</strong>
        <table class="suprlevel-print">
            <tr>
            <td class="suprlevel-print" style="width: 100%;">
              <%
              For Each key In ppe.keys
              %>
              <div style="float:left;padding-right:5px" align="center">
                    <image width="100px" class="ppeClass" id="<%=key%>"  src="images/<%=ppe.item(key)%>"/><br/>

              </div>
              <%
              Next
              %>
            </td>
            </tr>
            </table>
        </td>
     </tr>

     <tr>
        <td colspan = "4">
        <br/>
        <strong>Emergency Equipment required for this activity</strong>
        <table class="suprlevel-print">
            <tr>
            <td class="suprlevel-print" style="width: 100%;">
             <%
               For Each key In eq.keys
               %>
               <div style="float:left;padding-right:5px" align="center">
                    <image width="90px"  class="eqClass" id="<%=key%>" src="images/<%=eq.item(key)%>"/><br/>

               </div>
               <%
               Next
               %>
            </td>
            </tr>
            </table>
        </td>
     </tr>
<script type="text/javascript">

    str1 = '<%=strppe%>';
    str2 = '<%=streq%>';

    if(str1 != ''){
        var ppeItems =  JSON.parse(str1);
    }
    else{
        var ppeItems = [];
    }
    if(str2 != ''){
        var eqItems =  JSON.parse(str2);
    }
    else{
        var eqItems = [];
    }

    $('.ppeClass').each(function(){

        if($.inArray($(this).attr('id'), ppeItems)!== -1){
            $(this).show();
        }
        else{
            $(this).hide();
        }

     });
     $('.eqClass').each(function(){
             if($.inArray($(this).attr('id'), eqItems)!== -1){
                 $(this).show();
             }
             else{
                 $(this).hide();
             }

      });
</script>



<tr>
	<td colspan = "4">
	<br/>
	<strong>Work Activity Steps</strong>
		<table class="suprlevel-print" >
			<tr>
				<td style="width: 100%;">
					<%= Escape(strJobSteps)%>
					<br/>
				</td>
			</tr>
		</table>

	</td>
</tr>
	
	<tr>
	<td colspan = "4">	
				<tr>
					<td colspan = "3" class= "small-font">This SWMS has been developed through consultation with our employees and has been read, understood and signed by all employees undertaking the works.</td>
				</tr>
				<tr>
					<td>Print Names:</td><td>Signatures:</td><td>Dates:</td>
				</tr>
				<tr>
					<td><br/></td> <td><br/></td> <td><br/></td>
				</tr>
				<tr>
					<td><br/></td> <td><br/></td> <td><br/></td>
				</tr>
				<tr>
					<td><br/></td> <td><br/></td> <td><br/></td>
				</tr>
					<tr>
					<td><br/></td> <td><br/></td> <td><br/></td>
				</tr>
					<tr>
					<td><br/></td> <td><br/></td> <td><br/></td>
				</tr>
				<tr>
					<td colspan = "3" class= "small-font">Any changes, additions or deletions made to this SWMS are to be covered at a Tool Box meeting with all employees undertaking the works.</td>
				</tr>
	</td>
	</tr>


	<tr>
		<td colspan = "3"><strong>Supervisor: <%=strSupervisor%></strong></td>
	</tr>
	<tr>
		<td colspan = "3"><strong>Signature of supervisor:</strong>&nbsp;&nbsp;&nbsp;__________________________&nbsp;&nbsp;&nbsp;<strong>Date:&nbsp;&nbsp;&nbsp;</strong> __________________ </td>
	</tr>

</table>

</body>
</html>


