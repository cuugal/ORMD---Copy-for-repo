
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

Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")           
Escape = strReturn
End Function
%>


<script type="text/javascript" src="sorttable.js"></script>




<div id="wrapper">
 <div id="content">
    
<center>

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
numBuildingId = cint(Request.form("hdnBuildingId"))
numCampusId = cint(Request.form("hdnCampusId"))
strSupervisor = Request.form("hdnSuperV")
numFacultyId = cint(Request.form("hdnFacultyId"))
numFacilityId = cint(Request.form("hdnFacilityId"))

strHTask = Session("hdnHTask") 
numBuildingId =  cint("0"&Session("hdnBuildingId"))
numCampusId = cint("0"&Session("hdnCampusId"))
numFacultyId = cint("0"&Session("hdnFacultyId"))
numFacilityId = cint("0"&Session("hdnFacilityId"))
strSupervisor =Session("hdnSupervisor")
strOperation = Session("hdnOperationID")

searchType = session("searchType")

if strSupervisor = "0" then
 ' response.write("exception caught")
  strSupervisor = NULL 
end if  
%>

<%

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

' This query is used to collect the data to be displayed on the screen

if(searchType = "supervisor") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblFacilitySupervisor.numFacultyId, tblQORA.strSupervisor, tblQORA.strtaskDescription, tblRiskLevel.numGrade "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacility, tblFacilitySupervisor"
	strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
	if Len(strSupervisor) > 0 then
		strSQL = strSQL+" and tblFacilitySupervisor.strLoginID = '"& strSupervisor &"'"
	end if
	strSQL = strSQL+" and tblFacility.numFacilityID = tblQORA.numFacilityID"
	strSQL = strSQL+" and tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	'strSQL = strSQL + " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
	
	strSQL = strSQL+" union "
	
	strSQL = strSQL+"Select distinct(tblQORA.numQORAId) as numQORAId, tblFacilitySupervisor.numFacultyId, tblQORA.strSupervisor, tblQORA.strtaskDescription, tblRiskLevel.numGrade "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblOperations , tblFacilitySupervisor "
	strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
	if Len(strSupervisor) > 0 then
		strSQL = strSQL+" and tblFacilitySupervisor.strLoginID = '"& strSupervisor &"'"
	end if
	strSQL = strSQL+" and tblOperations.numOperationID = tblQORA.numOperationID"
	strSQL = strSQL+" and tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL + " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "

	'also, 'user' level audits
	strSQL = strSQL+" union "

    	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
        	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacilitySupervisor"

        	strSQL = strSQL+" Where tblFacilitySupervisor.strLoginId = tblQORA.strSupervisor "
        	strSQL = strSQL+" and tblFacilitySupervisor.numFacultyID = "&numFacultyId

            strSQL = strSQL+" and tblqora.numoperationid = 0 and tblqora.numfacilityid = 0 "
        	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"

        	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "

end if

if(searchType = "location") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacility, tblBuilding, tblCampus, tblFacilitySupervisor "
	'strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "&numFacultyId
	
	strSQL = strSQL+" Where tblQORA.numFacilityID = tblFacility.numFacilityID"
	strSQL = strSQL+" and tblBuilding.numBuildingID = tblFacility.numBuildingID"	
	strSQL = strSQL+" and tblCampus.numCampusID = tblBuilding.numCampusID"
	strSQL = strSQL+" and tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"

    if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
		strSQL = strSQL+" and tblFacilitySupervisor.numFacultyID = "&numFacultyId
	end if
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

	strSQL = strSQL+" Where tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorId"
	strSQL = strSQL+" and tblQORA.numOperationID = tblOperations.numOperationID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
    if Len(strOperation) > 0 and (strOperation <> 0) then
		strSQL = strSQL+" and tblOperations.numOperationID = "&strOperation
	end if
    if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
		strSQL = strSQL+" and tblFacilitySupervisor.numFacultyID = "&numFacultyId
	end if
	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
   ' Response.write(strSQL)
end If


if(searchType = "user") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacilitySupervisor"

	strSQL = strSQL+" Where tblFacilitySupervisor.strLoginId = tblQORA.strSupervisor "
	if Len(strSupervisor) > 0 then
    		strSQL = strSQL+" and tblFacilitySupervisor.strLoginID = '"& strSupervisor &"'"
    	end if
	strSQL = strSQL+" and tblFacilitySupervisor.numFacultyID = "&numFacultyId

    strSQL = strSQL+" and tblqora.numoperationid = 0 and tblqora.numfacilityid = 0 "
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"

	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
    'Response.write(strSQL)
end If


' added 3may18 DLJ
if(searchType = "templates") Then
strOperation = 75
numFacultyId = 28
'these two values correspond to david lloyd-jones TEMPLATES

	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblOperations, tblFacilitySupervisor"

	strSQL = strSQL+" Where tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorId"
	strSQL = strSQL+" and tblQORA.numOperationID = tblOperations.numOperationID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL+" and tblOperations.numOperationID = "&strOperation
	strSQL = strSQL+" and tblFacilitySupervisor.numFacultyID = "&numFacultyId
	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
   ' Response.write(strSQL)
end if








if(searchType = "task") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblFacilitySupervisor.numFacultyId, tblQORA.strSupervisor, tblQORA.strtaskDescription, tblRiskLevel.numGrade "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacility, tblFacilitySupervisor"
	
	where = 0
	if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
		strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
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
	strSQL = strSQL+" and tblQORA.numFacilityID > 0"
	strSQL = strSQL+" and tblFacility.numFacilityID = tblQORA.numFacilityID"
	strSQL = strSQL+" and tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"

	strSQL = strSQL+" union "

	strSQL = strSQL+"Select distinct(tblQORA.numQORAId) as numQORAId, tblFacilitySupervisor.numFacultyId, tblQORA.strSupervisor, tblQORA.strtaskDescription, tblRiskLevel.numGrade "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblOperations , tblFacilitySupervisor "

	where = 0
	if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
		strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
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
	strSQL = strSQL+" and tblQORA.numOperationID > 0"
	strSQL = strSQL+" and tblOperations.numOperationID = tblQORA.numOperationID"
	strSQL = strSQL+" and tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL + " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "


	' Third case where there is no linked operation or faculty
	strSQL = strSQL+" union "

    strSQL = strSQL+"Select distinct(tblQORA.numQORAId) as numQORAId, tblFacilitySupervisor.numFacultyId, tblQORA.strSupervisor, tblQORA.strtaskDescription, tblRiskLevel.numGrade "
    strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacilitySupervisor "

    where = 0
    if Len(numFacultyId) > 0 and (numFacultyId <> 0) then
        strSQL = strSQL+" Where tblFacilitySupervisor.numfacultyId = "& numFacultyId
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
    strSQL = strSQL+" and tblQORA.numOperationID = 0 and tblQORA.numFacilityID = 0"
    strSQL = strSQL+" and tblQORA.strSupervisor = tblFacilitySupervisor.strLoginId"
    strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	'response.write(strSQL)
	
end if

if(searchType = "myfac") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblFacility, tblBuilding, tblCampus, tblFacilitySupervisor "
	
	
	strSQL = strSQL+" Where tblQORA.numFacilityID = tblFacility.numFacilityID"
	strSQL = strSQL+" and tblBuilding.numBuildingID = tblFacility.numBuildingID"	
	strSQL = strSQL+" and tblCampus.numCampusID = tblBuilding.numCampusID"
	strSQL = strSQL+" and tblFacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"

    if cint(numFacultyId) <> -1 then
        strSQL = strSQL+" and tblFacilitySupervisor.numfacultyId = "&numFacultyId
    end if

	if Len(numFacilityId) > 0 and (numFacilityID <> 0) then
		strSQL = strSQL+" and tblFacility.numFacilityID = "&numFacilityId
	end if

	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"
	strSQL = strSQL + " Order by tblQORA.numFacilityID,tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
	
end if

    if(searchType = "myop") then
	strSQL = "Select distinct(tblQORA.numQORAId) as numQORAId, tblQORA.*, tblRiskLevel.* "
	strSQL = strSQL+" from tblQORA, tblRiskLevel, tblOperations, tblFacilitySupervisor"
	
	strSQL = strSQL+" WHERE tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorId"
	strSQL = strSQL+" and tblQORA.numOperationID = tblOperations.numOperationID"
	strSQL = strSQL+" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel"

    if cint(numFacultyId) <> -1 then
        strSQL = strSQL+" and tblFacilitySupervisor.numfacultyId = "&numFacultyId
    end if

	if Len(strOperation) > 0 and (strOperation <> 0) then
		strSQL = strSQL+" and tblOperations.numOperationID = "&strOperation
	end if


	strSQL = strSQL+ " Order by tblQORA.strSupervisor, tblRiskLevel.numGrade, tblQORA.strTaskDescription "
	

end if


'response.write strSQL
'response.end

'**********************Fire the Query**************************************************
'***********************establishing the database connection***************
set connSearch = Server.CreateObject("ADODB.Connection")
connSearch.open constr

'*************************Defining the recordset***************************
set rsSearch = Server.CreateObject("ADODB.Recordset")
rsSearch.Open strSQL, connSearch, 3, 3 

if rsSearch.EOF = TRUE then %>
Sorry, no records are present under this selection. Please try again....
<%else
dim SQLInsert

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr



while not rsSearch.Eof 

	'Code to fix apostrophe problem. 
	'Dont know why it was ever coded to insert and then retrieve but thats how it has been done.
 	
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

 
   SQLInsert ="Insert into tblQORATemp(numQORAId,numFacultyId,strSupervisor,strTaskDescription ) Values "_
   &" ("& rsSearch("numQORAId")  &","_
   &" "& rsSearch("numFacultyId")  &","_
   &" '"& strSupervisor &"',"_
   &" '"& strTaskDescription &"')"
   
   '&" "& rsSearch("numFacilityId")  &","_
   'Response.write(SQLInsert)
set rsTest = Server.CreateObject("ADODB.Recordset")
rsTest.Open SQLInsert, conn, 3, 3 
  
  rsSearch.Movenext
 ' i = i + 1
wend

'%><%'/*/*/*/*/*/*/*/*/*/*/*/*/* IMPORTANT */*/*/*/*/*/*/*/*/*/*response.write(strSQL)
i = 0
 
 
'*************************Defining the recordset*******************************************************
strSQL = "SELECT distinct(tblQORATemp.numQORAId) as numQORAId, tblQORA.numFacultyId, tblQORA.numFacilityId, tblQORA.strSupervisor, "_
 &" strFacultyName, strRoomName,strRoomNumber, tblQORA.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText,strCampusName,strBuildingName, null as numOperationID, null as strOperationName, strDateActionsCompleted, "_
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
 
 &"SELECT distinct(tblQORATemp.numQORAId) as numQORAId, tblQORA.numFacultyId, tblQORA.numFacilityId, tblQORA.strSupervisor, "_
 &" strFacultyName, null as strRoomName, null as strRoomNumber, tblQORA.strTaskDescription, "_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText, null as strCampusName, null as strBuildingName, tblQORA.numOperationID, strOperationName, strDateActionsCompleted, "_
 &" strGivenName, strSurname "_
 
 &" FROM tblQoraTemp, tblFaculty, tblOperations ,tblQORA, tblRiskLevel,tblFacilitySupervisor "_
 
 &" Where tblQORATemp.numQoraID = tblQORA.numQORAId "_
 
 &" and tblQORA.numOperationID = tblOperations.numOperationId "_
 
 &" and tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID "_
 &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_

 &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
 &" order by numFacilityID, numOperationID "_

  &" union all "_

  &"SELECT distinct(tblQORATemp.numQORAId) as numQORAId, tblQORA.numFacultyId, tblQORA.numFacilityId, tblQORA.strSupervisor, "_
  &" strFacultyName, null as strRoomName, null as strRoomNumber, tblQORA.strTaskDescription, "_
  &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
  &" boolFurtherActionsChemicalRA, dtReview, boolSWMSRequired,"_
  &" boolFurtherActionsGeneralRA,dtDateCreated,strText, null as strCampusName, null as strBuildingName, null as numOperationID, null as strOperationName, strDateActionsCompleted, "_
  &" strGivenName, strSurname "_

  &" FROM tblQoraTemp, tblFaculty ,tblQORA, tblRiskLevel, tblFacilitySupervisor "_

  &" Where tblQORATemp.numQoraID = tblQORA.numQORAId "_
  &" and tblQORA.numOperationId = 0 and tblQORA.numFacilityID = 0 "_

  &" and tblQORA.strSupervisor = tblFacilitySupervisor.strLoginId "_
  &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_

  &" and tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "_
  &" order by strSupervisor"

  set rsFaculty = Server.CreateObject("ADODB.Recordset")
  rsFaculty.Open strSQL, conn, 3, 3 
'*******************************************************************************************************


  dim tFac
  dim first_time
  if not rsFaculty.EOF then
  	tFac = rsFaculty("numFacultyId")
  	tFaci = rsFaculty("numFacilityId")
  	tOper = rsFaculty("numOperationID")
  	tUser = rsFaculty("strSupervisor")
  else 
  		%>Sorry, no records are present under this selection. Please try again....<%
  end if  
  first_time=true 
  rowID = 1%>

  <%  
  while not rsFaculty.EOF		
 	
  	if (searchType = "user" and tUser <> rsFaculty("strSupervisor")) or tFaci <> rsFaculty("numFacilityId") or tOper <> rsFaculty("numOperationID") or first_time then%>
  <%'************************Change format of header when switching faculties *****************%>
    
   
   
   <% if(rsFaculty("strRoomNumber") <> "") then %>
  	<table width="100%" class="searchResultsFromMenu" id="id1<%=rowID%>">
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
  			<strong>Assessor: </strong><%=rsFaculty("strGivenName")%>&nbsp;<%=rsFaculty("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="3"><strong>Faculty: </strong><%=rsFaculty("strFacultyName")%>&nbsp;&nbsp;&nbsp;
  			</td>
            <%
            set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select count(numQORAId) as numRA, sum(iif(dtReview > Date() , 1 , 0 )) as numCurrent, strRoomName from tblFacility, tblQORA "_
                &" where tblFacility.numFacilityId = "&rsFaculty("numFacilityId")_
                &" and tblQORA.numFacilityId = tblFacility.numFacilityId"_
                &" group by strRoomName"

            'response.write strControls
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
 	
                %>	
              <td class="campus" colspan="3"><strong>Current Risk Assessments: </strong><%=rsControls("numCurrent")%>/<%=rsControls("numRA")%></td>
  		<tr>
  	</table>
 
  <%elseif(rsFaculty("strOperationName") <> "") then %>
  	<table width="100%" class="searchResultsFromMenu" id="id1<%=rowID%>">
  	<tr>
  			<td class="campus">
  			<strong>Assessor: </strong><%=rsFaculty("strGivenName")%>&nbsp;<%=rsFaculty("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
  			<td class="campus" colspan="3"><strong>Operation: </strong><%=rsFaculty("strOperationName")%>&nbsp;&nbsp;&nbsp;
  			</td>	
             <%
            set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select count(numQORAId) as numRA, sum(iif(dtReview > Date() , 1 , 0 )) as numCurrent, strOperationName from tblOperations, tblQORA "_
                &" where tblOperations.numOperationId = "&rsFaculty("numOperationId")_
                &" and tblQORA.numOperationId = tblOperations.numOperationId"_
                &" group by strOperationName"

            'response.write strControls
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
 	
                %>	
              <td class="campus" colspan="3"><strong>Current Risk Assessments: </strong><%=rsControls("numCurrent")%>/<%=rsControls("numRA")%></td>
  			
  		<tr>
  	</table>
   

   


    <% else %>
    	<table width="100%" class="searchResultsFromMenu usermenu" id="id1<%=rowID%>">
      	<tr>
      			<td class="campus">
      			<strong>Assessor: </strong><%=rsFaculty("strGivenName")%>&nbsp;<%=rsFaculty("strSurname")%>&nbsp;&nbsp;&nbsp;</td>
      			<td class="campus" colspan="3"><strong>Faculty: </strong><%=rsFaculty("strFacultyName")%>&nbsp;&nbsp;&nbsp;
      			</td>
                 <%
                set connControls = Server.CreateObject("ADODB.Connection")
      			connControls.open constr
    			' setting up the recordset
       			strControls ="Select count(numQORAId) as numRA, sum(iif(dtReview > Date() , 1 , 0 )) as numCurrent from tblFacilitySupervisor, tblQORA "_
                    &" where tblFacilitySupervisor.strLoginId = '"&rsFaculty("strSupervisor")&"'"_
                    &" and tblQORA.strSupervisor = tblFacilitySupervisor.strLoginId"_
                    &" group by strLoginId"

                'response.write strControls
      			set rsControls = Server.CreateObject("ADODB.Recordset")
            	rsControls.Open strControls, connControls, 3, 3

                    %>
                  <td class="campus" colspan="3"><strong>Current Risk Assessments: </strong><%=rsControls("numCurrent")%>/<%=rsControls("numRA")%></td>

      		<tr>
      	</table>
    <% end if %>

  		<% rowID = rowID +1 %>
  		<table width="100%" class="sortable searchResultsFromMenu" id="id1<%=rowID%>">
  		<tr>
            <% if session("LoggedIn") then %><th class="singleaction">&nbsp;</th><% end if %> 
  			<th class="qoraID">Ra No.</th>
    		<th class="haztaskresult">Work Activity</th>
    		<th class="assochazards">Hazards</th>
    		<th class="currentcontrols">Current Controls</th>
    		<th class="risklevel">Risk Level</th>
    		<th class="furtheraction">Proposed Controls</th>
    		<th class="renewaldate">Review Date</th>
    		<th class="swms">SWMS</th>
  		</tr>
  <%
   		tFaci = rsFaculty("numFacilityId")
   		tOper = rsFaculty("numOperationID")
   		tFac = rsFaculty("numFacultyId")
   		tUser = rsFaculty("strSupervisor")
   		first_time = false
   		rowID = rowID+1
   	 end if%>
  <% 
     date_d = day(rsFaculty("dtDateCreated"))
     date_m = month(rsFaculty("dtDateCreated"))
     date_y = Year(rsFaculty("dtDateCreated")) + 2
     dtRDate = cstr(date_d)+"/"+cstr(date_m)+"/"+ cstr(date_y)
     %>
  	<tr>
         <% if session("LoggedIn") then %> <td><a href="#" data-toggle="modal" data-target="#CopyModal" data-qora="<%=rsFaculty("numQORAID")%>">Copy</a></td> <% end if %>
    	<td><%=Escape(rsFaculty("numQORAId"))%></td>
    	<!--td><a target="Operation" title="Click to edit this Risk Assessment." href="EditQORA.asp?numCQORAId=<%=rsFaculty("numQORAId")%>"><%=rsFaculty("strTaskDescription")		%></td-->
		<td><%=rsFaculty("strTaskDescription") %></td>
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
    	<td><center><%=Escape(rsFaculty("strAssessRisk"))%></center></td>
    	
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
   <!-- <td><center><%=dtRDate%></center></td> -->
<%
              dim today
         today = Date()
               %>
   <td <% If not isNull(rsFaculty("dtReview")) and DateDiff("d", rsFaculty("dtReview"), today) > 1 Then %>style="color:red;font-weight:bold" <%end if %> ><center>  <%=rsFaculty("dtReview")%></center></td>
   
   
   <td><center>
   <% If rsFaculty("boolSWMSRequired") = true Then %>
       <form method="post" action="SWMS.asp">
         <input type="submit" value="SWMS" name="btnSWMS" />
         <input type="hidden" name="hdnQORAId" value="<%=rsFaculty("numQORAId")%>" />
         <input type="hidden" name="hdnNoSaveBeforeSWMS" value="nosave"/>
       </form>

  <% End if%>
   
   </center></a></td>
   
   
  </tr>
  <% 
    if (searchType = "user" and tUser <> rsFaculty("strSupervisor")) or tFaci <> rsFaculty("numFacilityId") or tOper <> rsFaculty("numOperationID") or first_time then
   		%> </table>  <%  
    end if     
 rsFaculty.Movenext
 wend
    %>

</table>

</div>
<!-- close content -->
</div>
<!-- close wrapper -->

<%
set rsClear = Server.CreateObject("ADODB.Recordset")
rsClear.Open "delete from tblQORATemp", conn, 3, 3 
end if
 if session("LoggedIn") then 

    Dim connFaci
		Dim rsFillFaci
		Dim strSQLFaci
	  
		'Database Connectivity Code 
		set connFaci = Server.CreateObject("ADODB.Connection")
		connFaci.open constr
	   
		' setting up the recordset
	   
		strSQLFaci ="Select * from tblFacility "

		if session("numSupervisorId") <> 1 then
			strSQLFaci =strSQLFaci&" WHERE numFacilitySupervisorId = "& session("numSupervisorId")
		end if
		strSQLFaci = strSQLFaci&" order by strRoomNumber"
		set rsFillFaci = Server.CreateObject("ADODB.Recordset")
		rsFillFaci.Open strSQLFaci, connFaci, 3, 3

    Dim connProj
		Dim rsFillProj
		Dim strSQLProj
	  
		'Database Connectivity Code 
		set connProj = Server.CreateObject("ADODB.Connection")
		connProj.open constr
	   
		' setting up the recordset
	   
		strSQLProj ="Select * from tblOperations "

		if session("numSupervisorId") <> 1 then
			strSQLProj =strSQLProj&" WHERE numFacilitySupervisorId = "& session("numSupervisorId")
		end if
		strSQLProj = strSQLProj&" order by strOperationName"
		set rsFillProj = Server.CreateObject("ADODB.Recordset")
		rsFillProj.Open strSQLProj, connProj, 3, 3
        

%>


<div class="modal fade" id="CopyModal" tabindex="-1" role="dialog" aria-labelledby="CopyModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
          <h4 class="modal-title" id="exampleModalLabel">New message</h4>
      </div>
      <div class="modal-body">
        <form id="copyForm">
          <input type="hidden" class="form-control" id="qora" name="qora"/>
            <input type="hidden" name="mode" id="searchType" value="user"/>
          <div class="form-group">
            <label for="recipient-name" class="control-label">To Facility:</label>
           <select class="form-control" autocomplete="off" id="myfacility" size="1" name="cboFacility" tabindex="1" onchange="$('#searchType').val('location');$('#submitCopy').html('Copy to Location');">
			<option value="0">Select any one</option>
			<%
				while not rsFillFaci.Eof
					concat = rsFillFaci("strRoomNumber")&" "&rsFillFaci("strRoomName")
						%>   
					<option value="<%=rsFillFaci("numFacilityId")%>"><%=concat%></option>
			<% 
				rsFillFaci.Movenext	
				wend 
			%>
			</select>  
          </div>
            <hr/>
            <b>OR</b>
            <hr />
           <div class="form-group">
            <label for="recipient-name" class="control-label">To Operation:</label>
           <select class="form-control" autocomplete="off" id="myoperation" name="cboOperation" id="cboOperation" Onchange="$('#searchType').val('operation');$('#submitCopy').html('Copy to Operation');">
			<option value="0">Select any one</option>
			<%
				while not rsFillProj.Eof
						%>   
					<option value="<%=rsFillProj("numOperationId")%>">
						<%=rsFillProj("strOperationName")%></option>
			<% 
				rsFillProj.Movenext	
				wend 
			%>
			</select>  
          </div>
          <hr/>
                      <b>OR</b>
                      <hr />
                      <b>Copy to My General Risk Assessments</b> (leave the above fields blank)
        </form>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
        <button type="button" id="submitCopy" class="btn btn-primary">Copy</button>
      </div>
    </div>
  </div>
</div>

<script type="text/javascript">
    $('#CopyModal').on('show.bs.modal', function (event) {
        var button = $(event.relatedTarget); // Button that triggered the modal
        var qora = button.data('qora'); // Extract info from data-* attributes
        // If necessary, you could initiate an AJAX request here (and then do the updating in a callback).
        // Update the modal's content. We'll use jQuery here, but you could use a data binding library or other methods instead.
        var modal = $(this);
        modal.find('.modal-title').text('Copy Risk Assessment: ' + qora);
        modal.find('.modal-body #qora').val(qora);
    });

    $(function () {
        //twitter bootstrap script
        $("#submitCopy").click(function () {
            $.ajax({
                type: "POST",
                url: "AJAXCopy.asp",
                data: $('#copyForm').serialize(),
                success: function (data) {
                    var obj = jQuery.parseJSON(data);
                    var newRA = obj.result;
                    alert("Copied to RA " + newRA);
                    $("#CopyModal").modal('hide');
                    $("#refreshResults").submit();
                },
                error: function () {
                    alert("failure");
                }
            });
        });
    });

    $('#ArchiveModal').on('show.bs.modal', function (event) {
        var button = $(event.relatedTarget) // Button that triggered the modal
        var qora = button.data('qora') // Extract info from data-* attributes
        // If necessary, you could initiate an AJAX request here (and then do the updating in a callback).
        // Update the modal's content. We'll use jQuery here, but you could use a data binding library or other methods instead.
        var modal = $(this)
        modal.find('.modal-title').text('Archive Risk Assessment: ' + qora)
        modal.find('.modal-body #archiveQora').val(qora)
    });



    $(function () {
        //twitter bootstrap script
        $("#submitArchive").click(function () {
            $.ajax({
                type: "POST",
                url: "AJAXArchive.asp",
                data: $('#archiveForm').serialize(),
                success: function (data) {
                    var obj = jQuery.parseJSON(data);
                    var newRA = obj.result;
                    alert("Archived RA " + newRA);
                    $("#ArchiveModal").modal('hide');
                    $("#refreshResults").submit();
                },
                error: function () {
                    alert("failure");
                }
            });
        });
    });

    </script>

<% end if %>