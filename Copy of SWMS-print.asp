<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!-- deleted by DLJ %
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
dim loginId
loginId = session("strLoginId")
testval = request.form("hdnQORAID")
'Response.Write(loginId)
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
	<h1>Safe Work Method Statement (SWMS)</h1>
	</td>
	<td align="right"> <h4> Risk Assessment Number <%=testval%></h4></td>
	</tr>

     	

 <%

	testval = request.form("hdnQORAID")
	'Response.write(testval)
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
	strAssessRisk = rsSearch("strAssessRisk")
	strControlRiskDesc = rsSearch("strControlRiskDesc")
	strText = rsSearch("strText")
	strJobSteps = rsSearch("strJobSteps")
	dtDateCreated = rsSearch("dtDateCreated")
	strInherentRisk = rsSearch("strInherentRisk")
	strDate = rsSearch("strDateActionsCompleted")
	boolswms = rsSearch("boolFurtherActionsSWMS")
	boolCRA = rsSearch("boolFurtherActionsChemicalRA")
	boolGRA = rsSearch("boolFurtherActionsGeneralRA")
	
	    
  

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

	strFacultyName = rsF("strFacultyName") 
	
	'code for Facility Name
	set connFacility = Server.CreateObject("ADODB.Connection")
	connFacility.open constr
	' setting up the recordset
	strSQL ="Select * from tblFacility where numFacilityId ="& numFacilityId
	set rsFillFacility = Server.CreateObject("ADODB.Recordset")
	rsFillFacility.Open strSQL, connFacility, 3, 3
	  
	strRoomName = rsFillFacility("strRoomName")
	strRoomNo = rsFillFacility("strRoomNumber")    
	'strFacilityName =  cstr(strRoomNo) + " - " + cstr(strRoomName) 
%>
<tr>
      		<td colspan="2" align="left"><strong> Task: <%=rsSearch("strTaskDescription")%> </strong></td>
      		
      		
      		<% if strAssessRisk="L" then%> <td class="low" align="right" width="250"> Residual Risk Level: <strong>Low </strong></td><%end if%>
      		<% if strAssessRisk="M" then%> <td class="medium" align="right" width="250"> Residual Risk Level: <strong>Medium</strong> </td><%end if%>
      		<% if strAssessRisk="H" then%> <td class="high" align="right" width="250"> Residual Risk Level:<strong> High </strong></td><%end if%>
      		<% if strAssessRisk="E" then%> <td class="extreme" align="right" width="250"> Residual Risk Level:<strong> Extreme <strong></td><%end if%>
      	</tr>

	<tr>
 	<td colspan="3">

 <table class="suprlevel-print" style="width: 100%;">
 <tr>    		
    		<td class="campus" colspan="2"><strong>Campus: </strong><%=rsc("strCampusName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="2"><strong>Building: </strong><%=rsb("strBuildingName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td><strong>Room Name: </strong><%=strRoomName%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="2"><strong>Room Number: </strong><%=strRoomNo%>&nbsp;&nbsp;&nbsp;
    		</td>
  		</tr>
  		<tr>
  			<td class="campus" colspan="2">
  			<strong>Supervisor: </strong><%=strSupervisor%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="5"><strong>Faculty: </strong><%=strFacultyName%>&nbsp;&nbsp;&nbsp;
  			</td>		
  		</tr>
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
        	<Strong>Assessor:</strong><%=strAssessor%></td>
        	<td class="campus" colspan="2"><strong>Date Last Modified (dd/mm/yyyy):</strong>&nbsp;&nbsp;&nbsp;<%=dtDateCreated%></td>
          	<td class="campus" colspan="3"><Strong>Review Date:&nbsp;&nbsp;&nbsp;</strong><%=rsSearch("dtReview")%></td>
        </tr>
	</table>
	</td>
	</tr>

	<tr>
	<td colspan = "4">
	<h4>Hazards Relating to this task</h4>
	<table class="suprlevel-print" style="margin 0 auto; width:90%; margin-left:40px">
		<tr>
		<td style="width: 100%;">
		<strong>TASK HAZARDS: </strong><br/>
<%=Escape(strHazardsDesc)%><br/>

		</td>
		</tr>
		<tr>
		<td style="width: 100%;">
		<strong>INHERENT RISKS:</strong><br/>
<%=Escape(strInherentRisk)%><br/>
	
		</td>
		</tr>
		</table>
	</td>
	</tr>
	
	<tr>
	<td colspan = "4">
	<h4>List Safety Equipment, Training, Signage and Information</h4>
	<table style="margin 0 auto; width:90%; margin-left:40px">
		<tr>
		<td class="suprlevel-print" style="width: 100%;">

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
<%=Escape(strControlsImplemented)%><br/>

		</td>
		</tr>
		</table>
	</td>
	</tr>


<tr>
	<td colspan = "4">
	<h4>Job Steps</h4>
	<table class="suprlevel-print" style="margin 0 auto; width:90%; margin-left:40px">
		<tr>
		<td style="width: 100%;">
		
<%= Escape(strJobSteps)%>
<br/>

		</td>
		</tr>
		</table>
		<br/>
	</td>
	</tr>
	

</table>
<br/>
<div>
      <strong> &nbsp;&nbsp;&nbsp; Supervisor: </strong><%=session("strName")%>    &nbsp;&nbsp;&nbsp;  <strong>Signature of supervisor:_______________________   &nbsp;&nbsp;&nbsp;   Date: __________________ </strong></div>
</div>
</div>
</body>
</html>


