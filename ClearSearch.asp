<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->

<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Current entries for this facility</title>
</head>
<%

Function Escape(sString)

'Replace any Cr and Lf with <br>
strReturn = Replace(sString , vbCrLf, "<br>")
strReturn = Replace(strReturn , vbCr , "<br>")
strReturn = Replace(strReturn , vbLf , "<br>")           
Escape = strReturn
End Function

dim numFacilityId
dim conn
dim rsShow
dim rsShowHeader
dim strSQL
dim date_d
dim date_m
dim date_y
dim dtRdate

numFacilityId = request.form("hdnFacilityId")
numFacilityId= cint(numFacilityId)
'response.write(val)

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr

' SQL for required report
'AA jan 2010 fixes for relationship
'strSQL = "SELECT DISTINCT (strFacultyName), strBuildingName, strCampusName, strRoomName, strRoomNumber,strFacilitySupervisor"_
  strSQL = "SELECT DISTINCT (strFacultyName), strBuildingName, strCampusName, strRoomName, strRoomNumber,numFacilitySupervisorID"_
&" FROM tblQORA, tblCampus, tblBuilding, tblFacility, tblFaculty "_
&" WHERE tblQORA.numFacilityId=tblFacility.numFacilityId And tblQORA.numFacultyId=tblFaculty.numFacultyId And tblQORA.numCampusId=tblCampus.numCampusId And "_ 
&" tblQORA.numBuildingId=tblBuilding.numBuildingId And tblQORA.numFacilityId= "& numFacilityId &" "_
&" ORDER BY strFacultyName, strCampusName, strBuildingname, strRoomName "
  
 set rsShowHeader = Server.CreateObject("ADODB.Recordset")
   'response.write(strSQL)
   rsShowHeader.Open strSQL, conn, 3, 3  

dim supId 
dim supName
	'AA jan 2010 now interested in numFacilitySupervisorID not strFacilitySupervisor
   supId = rsShowHeader("numFacilitySupervisorID")
    set rsFs = Server.CreateObject("ADODB.Recordset")
    'AA jan 2010 now interested in numSupervisorID not strLoginID
    rsFs.Open "Select * from tblFacilitySupervisor where numSupervisorID = '"& supId &"'", conn, 3, 3  
   
   supName = cstr(rsFs("strGivenName"))+ " " + cstr(rsFs("strSurName"))
   
 strSQL ="Select strTaskDescription,strHazardsDesc,strControlRiskDesc,strAssessRisk, "_
 &" boolFurtherActionsSWMS,boolFurtherActionsChemicalRA,boolFurtherActionsGeneralRA, "_
 &" dtDateCreated,strText "_
 &" From tblQORA where numFacilityId ="& numFacilityId &" Order by strTaskDescription" 

 set rsShow = Server.CreateObject("ADODB.Recordset")
   'response.write(strSQL)
   rsShow.Open strSQL, conn, 3, 3  
   
%>

<body link="#000000" vlink="#000000" alink="#000000">

   <table border="1" cellspacing="1" width="100%" id="table1" bordercolor="#000000" style="border-collapse: collapse">
     
    <%  
     '*/*/*/*/*/*/*/*/*/*/*/*/*// IMPORTANT   response.write("exception caught !")%>
    <tr><td width="100%" colspan="6" bgcolor="#FFFFFF">
		<input border="0" src="utslogo.gif" name="I1" width="184" height="41" type="image">&nbsp;
		 Risk Assessment&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		Date Created : <%=Date%><br>
&nbsp;</td></tr> 
    <tr><td width="100%" colspan="6" bgcolor="#000000">
    <font face="Tahoma" size="2" color="#FFFFFF"><b>
    F</b>aculty<b> / U</b>nit<b>
    N</b>ame<b>
    : <%=rsShowHeader("strFacultyName")%><br>
&nbsp;</b></td></tr>
 
  <tr>
    <td width="59%" colspan="3" bgcolor="#FFFFFF"><b>
	Campus :<%=rsShowHeader("strCampusName")%></b></td>
    <td width="40%" colspan="3" bgcolor="#FFFFFF"> <b>
	Building : <%=rsShowHeader("strBuildingName")%></b></td>
  </tr>
 
  <tr>
    <td width="59%" colspan="3" bgcolor="#FFFFFF"><b>
    &nbsp;Facility Name/Room Number : <%=rsShowHeader("strRoomName")%> 
    / <%=rsShowHeader("strRoomNumber")%></b></td>
    <td width="40%" colspan="3" bgcolor="#FFFFFF"><b>
	Supervisor :<%=supName%> </b> </td>
  </tr>
  <tr>
      <td width="26%" align="center"><b>
    Hazardous Task</b></td>
    <td width="18%" align="center"><b>
    Associated Hazards</b></td>
    <td width="15%" align="center"><b>
    Current Controls</b></td>
    <td width="13%" align="center"><b>
    Risk Level</b></td>
    <td width="16%" align="center"><b>
    Further Action</b></td>
    <td width="11%" align="center"><b>
    Renewal Date</b></td>
  </tr>
           
  <% 
     while not rsShow.EOF 
     
        date_d = day(rsShow(7))
        date_m = month(rsShow(7))
        date_y = Year(rsShow(7)) + 5
        
        dtRDate = cstr(date_d)+"/"+cstr(date_m)+"/"+ cstr(date_y)
     %>   
   <tr>
        <td width="26%" bordercolor="#000000" bgcolor="#FFFFFF"><%=rsShow("strTaskDescription")%>&nbsp;</td>
     <td width="18%" bordercolor="#000000" align="left" valign="top">
     <%=escape(rsShow("strHazardsDesc"))%></td>
    <td width="15%" bordercolor="#000000" align="left" valign="top">
    <%=escape(rsShow("strControlRiskDesc"))%></td>
    <td width="13%" bordercolor="#000000" align="center"><BR><BR><%=rsShow("strAssessRisk")%><p align="center">&nbsp;</td>
    
    <td width="16%" bordercolor="#000000" bgcolor="#FFFFFF">
    <%if rsShow("boolFurtherActionsSWMS")= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc">Safe Work Method Statements</a> <%end if%><BR>
    <%if rsShow("boolFurtherActionsChemicalRA")= true then %><a target="_blank" href="http://www.ocid.uts.edu.au/">Chemical Risk Assessment</a> <%end if%><BR>    
    <%if rsShow("boolFurtherActionsGeneralRA")= true then %><a target="_blank" href="http://www.ehs.uts.edu.au/sections/level2/internal/generalriskmgt.doc">General Risk 
    Assessment</a> <%end if%><BR><BR>
    Comments:  <%=escape(rsShow("strText"))%>
     </td>
  <td width="9%" bordercolor="#000000" align="center"><BR><BR><%=dtRDate%><p align="center">&nbsp;</td>
       
  </tr>
  <% 
    rsShow.Movenext
  wend
    %> 
</table>

	<div align="right">
		<p align="center"><br>
&nbsp;</p>
		<table border="1" width="28%" id="table2" style="border-collapse: collapse" bordercolor="#000000">
			<tr>
				<td><br>
				<b>&nbsp;&nbsp; Further Action Completed<br>
				<br>
&nbsp;&nbsp; Signed : ______________<br>
				<br>
&nbsp;&nbsp; Date:___________<br>
&nbsp;</b></td>
			</tr>
		</table>
		<p align="center"><br>
&nbsp;</div>
	<p align="center"><br>
&nbsp;</p>
	<p align="center">&nbsp;</p>
	<p align="center"><br>
&nbsp;</p>
	<div align="right">
&nbsp;</div>

</body>

</html>