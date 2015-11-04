<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>
<%
dim loginVal
dim loginId
loginVal = session("strAccessLevel")
loginId = session("strLoginId")


if len(loginVal)<=0 then
'response.write("exception caught")
else
'response.write(loginVal)
'response.write(loginId)

end if


'10July2007 DLJ fixed function below same as in searchResultsFromMenu.asp - on advice of http://classicasp.aspfaq.com/general/how-do-i-prevent-invalid-use-of-null-errors.html
'fixed a null error
Function Escape(sString)
'Replace any Cr and Lf with <br />
    if len(sString) > 0 then 
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
    else 
        strReturn = "" 
    end if 
Escape = strReturn

End Function



%>

<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<!--<link rel="stylesheet" type="text/css" href="orr.css" media="screen,print">-->
<link rel="stylesheet" type="text/css" href="orrprint.css" media="screen,print">
<title>Results of Online Risk Register search</title>
<base target="Menu">
<style type="text/css">
<!--
.style2 {color: #FFFFFF}
-->
</style>

</head>
<body link="#000000" vlink="#000000" alink="#000000" topmargin="25">


<div id="wrapper">

<div id="content">

<!-- outside table -->
<table class="mainprintable">
<tr>
	<td>
	  <img src="utslogo.gif" width="184" height="41" title="" align="left">
	</td>
	<td>
	<h1>Risk Assessment Register</h1>
	</td>
	<td class="date">Date Printed: <%=date%></td>
</tr>
<tr>
 <td colspan="3">



<%
'*******************Declaring the variables********************************************
Dim strHTask
Dim numBuildingId
Dim numCampusId
Dim numFacultyId
Dim strSupervisor
Dim numFacilityId
Dim Input1
Dim Input2
Dim Input3
Dim Input4
Dim Input5
Dim Input6
Dim RsSearch
Dim ConnSearch
Dim strSQL
Dim flag
'*******************Fetching the inputs************************************************
strHTask = Request.form("txtHazardoustask")
numBuildingId = Request.form("hdnBuildingId")
numCampusId = Request.form("hdnCampusId")
strSupervisor = Request.form("hdnSuperV")
numFacultyId = Request.form("hdnFacultyId")
numFacilityId = Request.form("cboRoom")

strHTask = Session("hdnHTask") 
numBuildingId =  Session("hdnBuildingId")
numCampusId = Session("hdnCampusId")
numFacultyId = Session("hdnFacultyId")
numFacilityId =Session("hdnFacilityId")
strSupervisor =Session("hdnSupervisor")

strOperation = Session("hdnOperationID")

intSearchType = Session("intSearchType")

searchType = session("searchType")

if strSupervisor = "0" then
 ' response.write("exception caught")
  strSupervisor = NULL 
end if  
'******************Checking for valid inputs and populate the SQL**********************

dim i
dim flg
dim fc
dim f
dim b
dim c
dim s

fc = false
f = false
b = false
c = false
flg = false
s = false

i = 0
'********************************************************************************************************

if(searchType = "supervisor") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel "
	strSQL = strSQL+" Where tblQORA.numfacultyId = "& numFacultyId
	if Len(strSupervisor) > 0 then
		strSQL = strSQL+" and tblQORA.strSupervisor = '"& strSupervisor &"'"
	end if
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL + " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
	
end if

if(searchType = "location") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacility, tblBuilding, tblCampus "
	strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
	
	strSQL = strSQL+" and tblQORA.numFacilityID = tblFacility.numFacilityID"
	strSQL = strSQL+" and tblBuilding.numBuildingID = tblFacility.numBuildingID"	
	strSQL = strSQL+" and tblCampus.numCampusID = tblBuilding.numCampusID"

	if Len(numFacilityId) > 0 and (numFacilityID <> 0) then
		strSQL = strSQL+" and tblFacility.numFacilityID = "&numFacilityId
	end if
	if Len(numBuildingId) > 0 and (numBuildingID <>0 )then
		strSQL = strSQL+" and tblBuilding.numBuildingId = "&numBuildingId
	end if
	if Len(numCampusId) > 0 and (numCampusID<>0) then
		strSQL = strSQL+" and tblCampus.numCampusId = "&numCampusId
	end if
	
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL + " Order by tblQORA.numFacilityID,tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
	
end if

if(searchType = "operation") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblOperations, tblFacilitySupervisor"
	strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
	
	if Len(strOperation) > 0 and (strOperation <> 0) then
		strSQL = strSQL+" and tblOperations.numOperationID = "&strOperation
	end if
	strSQL = strSQL+" and tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorId"
	strSQL = strSQL+" and tblQORA.numOperationID = tblOperations.numOperationID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
end if

if(searchType = "task") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel "
	
	where = 0
	if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
		strSQL = strSQL+" Where tblQORA.numfacultyId = "& numFacultyId
		where = 1
	end if
	
	if Len(strHTask) >0 and strHTask <> " " and strHTask <> "*" then
	 	if(isNumeric(strHTask)) then
	 		if(where = 0) then 
	 			strSQL =  strSQL + " where tblQORA.numQORAId = "&strHTask
	 		else
	 			strSQL =  strSQL + " and tblQORA.numQORAId = "&strHTask
	 		end if	
	 	else
	 		if(where = 0) then
	    		strSQL =  strSQL + " where tblQORA.strTaskDescription like '%"& strHTask &"%'"
	    	else
	    		strSQL =  strSQL + " and tblQORA.strTaskDescription like '%"& strHTask &"%'"
	    	end if	
	  	end if
	end if
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
end if

'response.write(strSQL)
'response.End()
'**********************Fire the Query**************************************************
'***********************establishing the database connection***************
set connSearch = Server.CreateObject("ADODB.Connection")
connSearch.open constr

'*************************Defining the recordset***************************
set rsSearch = Server.CreateObject("ADODB.Recordset")
rsSearch.Open strSQL, connSearch, 3, 3 

if rsSearch.EOF = TRUE then %>
<font color="#660033"><b>
<%else
dim SQLInsert

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

while not rsSearch.Eof 
	strSupervisor = rsSearch("strSupervisor")
	temp = instr(1,strSupervisor,"'",vbTextCompare)
      if temp >= 1 then 
         strSupervisor = Replace(strSupervisor,"'","''",1)
      end if

 	strTaskDescription = rsSearch("strtaskDescription")
	temp = instr(1,strTaskDescription,"'",vbTextCompare)
      if temp >= 1 then 
         strTaskDescription = Replace(strTaskDescription,"'","''",1)
      end if

 
   SQLInsert ="Insert into tblQORATemp(numQORAId,numFacultyId,numFacilityId,strSupervisor,strTaskDescription ) Values "_
   &" ("& rsSearch("numQORAId")  &","_
   &" "& rsSearch("numFacultyId")  &","_

   &" "& rsSearch("numFacilityId")  &","_
   &" '"& strSupervisor &"',"_
   &" '"& strTaskDescription &"')"
   
set rsTest = Server.CreateObject("ADODB.Recordset")
rsTest.Open SQLInsert, conn, 3, 3 
  
  rsSearch.Movenext
wend

i = 0
 

'*************************Defining the recordset*******************************************************
strSQL = "SELECT distinct(tblQORATemp.numQORAId) as numQORAId, tblQORATemp.numFacultyId, tblQORATemp.numFacilityId, tblQORATemp.strSupervisor, "_
 &" strFacultyName, strRoomName,strRoomNumber,tblQORATemp.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText,strCampusName,strBuildingName, null as strOperationName, strDateActionsCompleted, "_
 &" strGivenName, strSurname "_
 
 &" FROM tblQoraTemp, tblFaculty, tblFacility,tblQORA,tblCampus,tblBuilding,tblRiskLevel,tblFacilitySupervisor "_
 
 &" Where tblQORATemp.numQoraID = tblQORA.numQORAId "_
 
 &" and tblQORA.numFacilityID = tblFacility.numFacilityID"_
 &" and tblBuilding.numBuildingID = tblFacility.numBuildingID"_
 &" and tblBuilding.numCampusID = tblCampus.numCampusID"_
 
 &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
 &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID"_
 
 &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
 
 &" union all "_
 
 &"SELECT distinct(tblQORATemp.numQORAId) as numQORAId, tblQORATemp.numFacultyId, tblQORATemp.numFacilityId, tblQORATemp.strSupervisor, "_
 &" strFacultyName, null as strRoomName, null as strRoomNumber, tblQORATemp.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText, null as strCampusName, null as strBuildingName, strOperationName, strDateActionsCompleted, "_
 &" strGivenName, strSurname "_
 
 &" FROM tblQoraTemp, tblFaculty, tblOperations ,tblQORA,tblRiskLevel,tblFacilitySupervisor "_
 
 &" Where tblQORATemp.numQoraID = tblQORA.numQORAId "_
 
 &" and tblQORA.numOperationID = tblOperations.numOperationId "_
 
 &" and tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID "_
 &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_

 &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
 &" order by numFacilityID, strOperationName"
 
  set rsFaculty = Server.CreateObject("ADODB.Recordset")
  rsFaculty.Open strSQL, conn, 3, 3 
'*******************************************************************************************************

'response.write strSQL
'response.end

%> </b>

<%'************table to display the main information.............%>
<%
  dim tFac
  dim first_time
  tFac = rsFaculty("numFacultyId")
  tFaci = rsFaculty("numFacilityId")  
  toper = rsFaculty("strOperationName") 
  first_time = true %> 

<!--inside table-->

<table class="suprlevel-print" id="id12" style="width: 100%;">

<%  while not rsFaculty.EOF		
    
   if tFaci <> rsFaculty("numFacilityId") or tOper <> rsFaculty("strOperationName") or first_time then%>  
     
 	<%'*************************'%>
 	<%if not first_time then%>
 	<tr>
		<td colspan="6">
		&nbsp;
		</td>
	</tr>
	
	<%end if 
	
	if(rsFaculty("strRoomNumber") <> "") then %>
  		<tr>    		
    		<td class="campus">
    		<strong>Campus: </strong><%=rsFaculty("strCampusName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="3"><strong>Building: </strong><%=rsFaculty("strBuildingName")%>&nbsp;&nbsp;&nbsp;
    		<strong>Room Name: </strong><%=rsFaculty("strRoomName")%>&nbsp;&nbsp;&nbsp;</td>
    		<td class="campus" colspan="3"><strong>Room Number: </strong><%=rsFaculty("strRoomNumber")%>&nbsp;&nbsp;&nbsp;
    		</td>
  		</tr>
  		<tr>
  			<td class="campus">
  			<strong>Supervisor: </strong><%=rsFaculty("strGivenName")%>&nbsp;<%=rsFaculty("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="6"><strong>Faculty: </strong><%=rsFaculty("strFacultyName")%>&nbsp;&nbsp;&nbsp;
  			</td>		
  		<tr>
 
  <%elseif(rsFaculty("strOperationName") <> "") then %>
  		<tr>
  			<td class="campus">
  			<strong>Supervisor: </strong><%=rsFaculty("strGivenName")%>&nbsp;<%=rsFaculty("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="6"><strong>Operation: </strong><%=rsFaculty("strOperationName")%>&nbsp;&nbsp;&nbsp;
  			</td>		
  		<tr>
   <%end if%>	


	<tr>
	 <th>RA No.</th>
	 <th>Task</th>
	 <th>Hazards</th>
	 <th>Current Controls</th>
	 <th>Risk Level</th>
	 <th>Proposed Controls</th>
	 <th>Review Date</th> 
	</tr>
  
   <%
   tFaci = rsFaculty("numFacilityId")
   toper = rsFaculty("strOperationName") 
   first_time = false
   end if%>
           

<tbody>
   <tr>
   <td><%=Escape(rsFaculty("numQORAId"))%></td>
    <td><%=rsFaculty("strTaskDescription")%></td>
    <td><%=Escape(rsFaculty("strHazardsDesc"))%></td>
    <td><%
          
          testval = rsFaculty("numQORAId")
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
    <td><center><%=rsFaculty("strAssessRisk")%></center></td>
    
    <%'****** Further action required*********%>
    <!-- old further action code<td><%=Escape(rsFaculty("strText"))%>
   <%if rsFaculty("boolFurtherActionsSWMS")= true then %><BR><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%>
    <%if rsFaculty("boolFurtherActionsChemicalRA")= true then %><BR><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a> <%end if%>    
    <%if rsFaculty("boolFurtherActionsGeneralRA")= true then %><BR><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
    Assessment</a> 
	<%end if%>
	</td> -->
	<td><%
          
          testval = rsFaculty("numQORAId")
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

	<td><center><%=rsFaculty("dtReview")%></center></td> 
  </tr>
</tbody>

  <% 
        ' rsFacility.moveNext
   ' wend
   
   ' tempFac = tempFac -1
  ' wend
   rsFaculty.Movenext
  wend
    %> 

</td>
</tr>
</table>


</div>

</body>
</html>
<%set rsClear = Server.CreateObject("ADODB.Recordset")
rsClear.Open "delete from tblQORATemp", conn, 3, 3 
end if%>