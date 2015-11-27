<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
dim loginId
loginId = session("strLoginId")
'Response.Write(loginId)
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen">
<!--<link rel="stylesheet" type="text/css" href="orrprint.css" media="print" />-->
<title>Online Risk Register - UTS Risk Assessments</title>
<script language="javaScript" type="text/javascript">
<!--
function framebreakout()
{
  if (top.location != location) {
    top.location.href = document.location.href ;
  }
}


//-->
</script>
<script type="text/javascript" src="sorttable.js"></script>
</head>
<%
'Campbells borrowed code to escape the output 15/6/2006
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
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
  dim numSupervisorId
  
  
       QORAtype = request.form("QORAtype")
	  session("QORAtype") = QORAtype
	  
	  numOperationID = request.form("cboOperation")
      session("cboOperation") = numOperationId
	  
    cboFacility = request.form("cboFacility")
	  session("cboFacility") = cboFacility
      
      numSupervisorId = session("numSupervisorId")
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  

  %>
<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
  <div id="content">
  <!-- Break out of frame --> 
  <form target="_blank" action="SupRDateModified-print.asp">
    <h1 class="pagetitle">Risk Assessment Search Results &nbsp;&nbsp;&nbsp;<input type="submit" value="Print preview" /></h1>    
  </form>

</div>
<%


 if(QORAtype = "location") then 

 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblRiskLevel ,tblFacilitySupervisor "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  
  &" tblQORA.numFacilityId = "& cboFacility &" and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel and "_ 
  &" tblFacilitySupervisor.numSupervisorId = "& numSupervisorID &" ORDER BY tblRiskLevel.numGrade, strTaskDescription"
 end if
 
 
 if(QORAtype = "operation") then
	 strSQL = "SELECT * FROM tblQORA, tblOperations, tblRiskLevel ,tblFacilitySupervisor "_
  &" WHERE tblQORA.numOperationId = tblOperations.numOperationId and "_
  &" tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID and"_
  
  &" tblQORA.numOperationId = "& numOperationID &" and "_
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel and "_ 
  &" tblFacilitySupervisor.numSupervisorId = "& numSupervisorID &" ORDER BY tblRiskLevel.numGrade, strTaskDescription"
 end if

      
    set rsSearchH = server.CreateObject("ADODB.Recordset")
    rsSearchH.Open strSQL, Conn, 3, 3 
    if rsSearchH.EOF <> true then 
%>
	
    
    <% if(QORAtype="location") then %>
<table class="suprreportheader">
	<tr>
		<th>Campus:</th><td><%=cstr(rsSearchH("strCampusName")) %></td>
		<th>Building:</th><td><%=cstr(rsSearchH("strBuildingName")) %></td>
		<th>Room Name:</th><td><%=cstr(rsSearchH("strRoomName"))%></td>
		<th>Room Number:</th><td><%=cstr(rsSearchH("strRoomNumber"))%></td>
  
    </tr>
    <tr>
    	<th>Supervisor:</th><td><%Response.Write (strName)%></td>
    	<th>Faculty:</th><td colspan="5"><%Response.Write(strFacultyName) %></td>
 </tr>
 </table>
<% end if
if (QORAtype = "operation") then %>
<table class="suprreportheader">
	<tr>
		<th>Supervisor: </th><td><%=strName%></td>
		<th>Operation: </th><td><%=rsSearchH("strOperationName")%></td>	
	</tr>
</table>
<% end if %>
    <%
  
   'Response.Write(strSQL) 
  if not rsSearchH.EOF then 
       %>
    <br />
    <table class="sortable suprlevel" id="id13">
     <caption>
      To sort any column, click on a table heading.  To edit a risk assessment, click on its title under the &quot;Task&quot; column.
      </caption>
      <thead>
        <tr>
        	<th class="qoraID">Ra No.</th>
          	<th class="haztaskresult">Task</th>
    		<th class="assochazards">Hazards</th>
    		<th class="currentcontrols">Current Controls</th>
    		<th class="risklevel">Risk Level</th>
    		<th class="furtheraction">Proposed Controls</th>
    		<th class="renewaldate">Review Date</th>
    		<th class="swms">SWMS</th>
        </tr>
      </thead>
    	<tbody>
      <%

	 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",2,rsSearchH("dtDateCreated"))
    
    %>
      
        <tr>
        <td><%=Escape(rsSearchH("numQORAId"))%></td>
          <td><a target="Operation" title="Click to edit this Risk Assessment." href="EditQORA.asp?numCQORAId=<%=rsSearchH("numQORAID")%>"><%=rsSearchH("strTaskDescription")%></td>
          <!--		<td><% Response.Write(rsSearchH(11))%></td> -->
          <td><%=Escape(rsSearchH("strHazardsDesc"))%></td>
          <td><%
          
          testval = rsSearchH("numQORAId")
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
          <td><center>
              <%=rsSearchH("strAssessRisk")%>
            </center></td>
         <!-- old 'further action required' code <td><% Response.Write(rsSearchH("strText"))%>
            <%if rsSearchH("boolFurtherActionsSWMS")= true then %>
            <BR>
            <a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc" title="Safe Work Method Statement (in Microsoft Word format, 47 Kb).">Safe Work Method Statement</a>
            <%end if%>
            <%if rsSearchH("boolFurtherActionsChemicalRA")= true then %>
            <BR />
            <a target="_blank" href="http://www.ocid.uts.edu.au/" title="Chemical risk assessment at OCID.">Chemical Risk Assessment</a>
            <%end if%>
            <%if rsSearchH("boolFurtherActionsGeneralRA")= true then %>
            <BR />Detailed Risk Assessment<%end if%></td>
          <td><% Response.Write(rsSearchH(17))%></td>-->
          <td><%
          ' New code to put in the unimplemented risk controls
          
          testval = rsSearchH("numQORAId")
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
          
      	<td><center><%=rsSearchH("dtReview")%></center></td>
         <td><center>
        <% If rsSearchH("boolSWMSRequired") = true Then %>
                 <form method="post" action="SWMS.asp">
         <input type="submit" value="SWMS" name="btnSWMS" />
         <input type="hidden" name="hdnQORAId" value="<%=rsSearchH("numQORAId")%>" />
         <input type="hidden" name="hdnNoSaveBeforeSWMS" value="nosave"/>
         </form>

        <% End if%>
                 </center></td>
            </tr>
        <%
    rsSearchH.MoveNext  
 wend

 %>
      </tbody>
    </table>
    <%
 'end if 
 end if %>
 
<%else%>
<p>There are currently no Risk Assessments for this facility or operation!</p>
<%end if%>
</div>
</body>
</html>
