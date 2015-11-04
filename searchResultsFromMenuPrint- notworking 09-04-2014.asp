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


Function Escape(sString)
'Replace any Cr and Lf with <br>
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
<title>Search Quick and Obvious Risk Assessment Results</title>
<base target="Menu">
</head>
<body link="#000000" vlink="#000000" alink="#000000" topmargin="25">
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

if strSupervisor = "0" then
 ' response.write("exception caught")
  strSupervisor = NULL 
end if  
'******************Checking for valid inputs and populate the SQL**********************
'Response.write("Task:")
'Response.write(strHTask)
'Response.write("Building:")
'Response.write(numBuildingId)%>

<%
'Response.write("Campus:")
'Response.write(numCampusId)%><%
'Response.write("Login:")
'Response.write(strSupervisor)%><% 
'Response.write("Faculty:")
'Response.write(numFacultyId)%><% 
'Response.write("Room:")
'Response.write(numFacilityId)
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
strSQL = "Select * from tblQORA"
'********************************************************************************************************
    if strHTask = " " or strHTask ="*" then
    	strSQL = "Select * from tblQORA"
    end if
'*********************************************************************************************************
    if Len(strHTask) >0 and i = 0 then
        strSQL =  strSQL + " where strTaskDescription like '%"& strHTask &"%'"
        i = 1
        flg = true
    end if
 '********************************************************************************************************	
 	if numFacultyId <> 0 then
 	 if  i > 0 then
    	 strSQL =  strSQL + " and numfacultyId = "& numFacultyId &""
     else
         strSQL =  strSQL + " Where numfacultyId = "& numFacultyId &""
	      i = 1
 	 end if
 	 flg = true
 	 fc = true
 	end if
'********************************************************************************************************* 	
 	if numFacilityId <> 0 then
 	 if  i >0 then
    	 strSQL =  strSQL + " and numfacilityId = "& numfacilityId &""
     else
         strSQL =  strSQL + " Where numfacilityId = "& numfacilityId &""
	     i = 1
 	 end if
 	 flg = true
 	 f = true
 	end if
'********************************************************************************************************* 	
 	if numCampusId <> 0 then
 	if  i > 0 then
    	 strSQL =  strSQL + " and numCampusId = "& numCampusId &""
     else
         strSQL =  strSQL + " Where numCampusId = "& numCampusId &""
	     i =  1
 	 end if
   flg = true
    c = true
    
 	end if
'********************************************************************************************************* 	
 	if numBuildingId <> 0 then
 	if  i >0 then
    	 strSQL =  strSQL + " and numBuildingId = "& numBuildingId &""
     else
         strSQL =  strSQL + " Where numBuildingId = "& numBuildingId &""
	     i =  1
 	 end if
flg = true
b = true
  	end if
'*********************************************************************************************************  
if len(strSupervisor)>0 then
 	if  i >0 then
    	 strSQL =  strSQL + " and strSupervisor = '"& strSupervisor &"'"
     else
         strSQL =  strSQL + " Where strSupervisor = '"& strSupervisor &"'"
	     i =  1
 	 end if
flg = true
s = true
 end if
'********************************************************************************************************* 	
strSQL = strSQL + " Order by strTaskDescription "
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

 
   SQLInsert ="Insert into tblQORATemp(numQORAId,numFacultyId,numCampusID,numBuildingId,numFacilityId,strSupervisor,strTaskDescription ) Values "_
   &" ("& rsSearch("numQORAId")  &","_
   &" "& rsSearch("numFacultyId")  &","_
   &" "& rsSearch("numCampusId")  &","_
   &" "& rsSearch("numBuildingId")  &","_
   &" "& rsSearch("numFacilityId")  &","_
   &" '"& rsSearch("strSupervisor") &"',"_
   &" '"& rsSearch("strtaskDescription") &"')"
   
set rsTest = Server.CreateObject("ADODB.Recordset")
rsTest.Open SQLInsert, conn, 3, 3 
  
  rsSearch.Movenext
 ' i = i + 1
wend

'response.Write(i)
'i = 0

'**************************************************


'response.write("Count :")
'response.write(i)
'%><%'/*/*/*/*/*/*/*/*/*/*/*/*/* IMPORTANT */*/*/*/*/*/*/*/*/*/*response.write(strSQL)
i = 0
 

'*************************Defining the recordset*******************************************************
 strSQL = "SELECT tblQORATemp.numQORAId, tblQORATemp.numFacultyId, tblQORATemp.numFacilityId, "_
 &" strFacultyName, strRoomName,strRoomNumber,tblQORATemp.strTaskDescription, "_
 &" strHazardsDesc, strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText,strCampusName,strBuildingName,strGivenName,strSurname"_
 
 &" FROM tblQoraTemp, tblFaculty, tblFacility,tblQORA,tblCampus,tblBuilding,tblFacilitySupervisor "_
 
 &" WHERE tblQoRaTemp.numFacultyId=tblFaculty.numFacultyId And "_
 &" tblQORATemp.numQORAId=tblQORA.numQORAId And"_
 &" tblQoRaTemp.numFacilityId=tblfacility.numFacilityId And"_
 &" tblQORATemp.numCampusId = tblCampus.numCampusId And "_
 &" tblQORATemp.numBuildingId = tblBuilding.numBuildingId And "_
 &" tblQORATemp.strSupervisor = tblFacilitySupervisor.strLoginID "_
 
 &" GROUP BY tblQORATemp.numQORAId, tblQoraTemp.numFacultyId, "_
 &" tblQORATemp.numFacilityId, strFacultyName, strRoomName ,strRoomNumber,tblQORATemp.strTaskDescription,"_
 &" strHazardsDesc ,strControlRiskDesc,strAssessRisk,boolFurtherActionsSWMS,"_
 &" boolFurtherActionsChemicalRA,"_
 &" boolFurtherActionsGeneralRA,dtDateCreated,strText,strCampusName,strBuildingName,strGivenName,strSurname"_
 
 &" ORDER BY strFacultyName,strCampusName,strRoomName"
 
 
  set rsFaculty = Server.CreateObject("ADODB.Recordset")
  rsFaculty.Open strSQL, conn, 3, 3 
'*******************************************************************************************************

%> </b></font>

<%'************table to display the...... main information.............%>
<%
  dim tFac
  tFac = rsFaculty(1)
  tFaci = rsFaculty(2)   %>
   <p><img border="0" src="utslogo.gif" width="184" height="41">&nbsp;&nbsp;&nbsp;
	<b><font size="5">&nbsp;Quick and Obvious Risk Assessment</font></b></p>
   <table border="1" cellspacing="1" width="100%" id="AutoNumber2" bordercolor="#000000" style="border-collapse: collapse">
     
     <tr>
    <td width="100%" colspan="6" bgcolor="#000000">
    <font face="Tahoma" size="2" color="#FFFFFF"><b>
    F</b>aculty<b> / U</b>nit<b>
    N</b>ame <b>
    : <%=rsFaculty("strFacultyName")%></b></font></td>
    </tr>
    
    <tr>
    <td width="100%" colspan="6" bgcolor="#FFFFFF"><b>
    <font size="2" face="Tahoma">Campus : <%=rsFaculty("strCampusName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Building : <%=rsFaculty("strBuildingName")%>&nbsp; </font></b></td>
  </tr> 
     
  <tr>
    <td width="99%" bgcolor="#FFFFFF" align="center" colspan="6">
	<p align="left"><B><font size="2" face="Tahoma">Facility Name/Room Number : <%=rsFaculty("strRoomName")%> / <%=rsFaculty("strRoomNumber")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
	Supervisor :  <%=cstr(rsFaculty("strGivenName")) + " " + cstr(rsFaculty("strSurName")) %></font></B></td>
  </tr>
    
  <tr>
    <td width="99%" bgcolor="#FFFFFF" align="center" colspan="6">&nbsp;</td>
  </tr>
    
  <tr>
    <td width="26%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Hazardous Task</font></b></td>
    <td width="17%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Associated Hazards</font></b></td>
    <td width="20%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Current Controls</font></b></td>
    <td width="9%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Risk Level</font></b></td>
    <td width="16%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Further Action</font></b></td>
    <td width="11%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Renewal Date</font></b></td>
  </tr>
    
<%  while not rsFaculty.EOF		
      if tFac <> rsFaculty(1) then
     '*/*/*/*/*/*/*/*/*/*/*/*/*// IMPORTANT   response.write("exception caught !")%>
    <tr><td width="100%" colspan="6" bgcolor="#FFFFFF" bordercolorlight="#000000">&nbsp;</td></tr> 
    <tr>
    <td width="100%" colspan="6" bgcolor="#000000">
    <font face="Tahoma" size="2" color="#FFFFFF"><b>
    F</b>aculty<b> / U</b>nit<b>
    N</b>ame<b>
    : <%=rsFaculty("strFacultyName")%></b></font></td>
    </tr>

    <%
    tFac = rsFaculty(1) 
 
     end if%>
   <%if tFaci <> rsFaculty(2) then%>  
     
 <tr>
    <td width="100%" colspan="6" bgcolor="#FFFFFF"><b>
    <font size="2" face="Tahoma">Campus : <%=rsFaculty("strCampusName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Building : <%=rsFaculty("strBuildingName")%></font></b></td>
  </tr>
     
  <tr>
    <td width="99%" bgcolor="#FFFFFF" align="center" colspan="6">
	<p align="left"><B> <font size="2" face="Tahoma">Facility Name/Room Number :  <%=rsFaculty("strRoomName")%> / <%=rsFaculty("strRoomNumber")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Supervisor : <%=cstr(rsFaculty("strGivenName")) + " " + cstr(rsFaculty("strSurName")) %></font></B></td>
  </tr>
      
  <tr>
    <td width="99%" bgcolor="#FFFFFF" align="center" colspan="6">&nbsp;</td>
  </tr>
      
  <tr>
    <td width="26%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Hazardous Task</font></b></td>
    <td width="17%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Associated Hazards</font></b></td>
    <td width="20%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Current Controls</font></b></td>
    <td width="9%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Risk Level</font></b></td>
    <td width="16%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Further Action</font></b></td>
    <td width="11%" bgcolor="#000000" align="center"><b>
    <font face="Tahoma" size="2" color="#FFFFFF">Renewal Date</font></b></td>
  </tr>
      
   <%
   tFaci = rsFaculty(2)
   end if%>

           
  <% 
    ' while tempFac >0
       
     '   while not rsFacility.EOF     
        date_d = day(rsFaculty(12))
        date_m = month(rsFaculty(12))
        date_y = Year(rsFaculty(12)) + 5
        
        dtRDate = cstr(date_d)+"/"+cstr(date_m)+"/"+ cstr(date_y)
     %>   
   <tr>
        <td width="26%" bordercolor="#000000" bgcolor="#FFFFFF"><font face="Tahoma" size="1"><%=rsFaculty(6)%></font>&nbsp;</td>
     <td width="17%" bordercolor="#000000" bgcolor="#FFFFFF" align="left" valign="top"><font face="Tahoma" size="1">
     <%=Escape(rsFaculty("strHazardsDesc"))%></font></td>
    <td width="20%" bordercolor="#000000" bgcolor="#FFFFFF" align="left" valign="top"><font face="Tahoma" size="1">
    <%=Escape(rsFaculty("strControlRiskDesc"))%></font></td>
    <td width="9%" bordercolor="#000000" bgcolor="#FFFFFF" align="center"><font face="Tahoma" size="1"><BR><BR><%=rsFaculty("strAssessRisk")%></font><p align="center">&nbsp;</td>
    
    <td width="16%" bordercolor="#000000" bgcolor="#FFFFFF"><font face="Tahoma" size="1">
    <%if rsFaculty("boolFurtherActionsSWMS")= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/SWMS.doc">Safe Work Method Statements</a> <%end if%><BR>
    <%if rsFaculty("boolFurtherActionsChemicalRA")= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a> <%end if%><BR>    
    <%if rsFaculty("boolFurtherActionsGeneralRA")= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
    Assessment</a> <%end if%><BR><BR>
    <b><u>Comments</u> </b>
     <BR>  <%=Escape(rsFaculty("strText"))%></font>
    </td>
  <td width="9%" bordercolor="#000000" bgcolor="#FFFFFF" align="center"><font face="Tahoma" size="1"><BR><BR><%=dtRDate%></font><p align="center">&nbsp;</td>
       
  </tr>
  <% 
        ' rsFacility.moveNext
   ' wend
   
   ' tempFac = tempFac -1
  ' wend
   rsFaculty.Movenext
  wend
    %> 
</table>
</body>
</html>
<%set rsClear = Server.CreateObject("ADODB.Recordset")
rsClear.Open "delete from tblQORATemp", conn, 3, 3 
end if%>