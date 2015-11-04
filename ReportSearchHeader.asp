<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<!-- Generic Metadata -->
<meta name="description" content="The Online Risk Register of the University of Technology, Sydney." />
<meta name="keywords" content="online,risk,register,environment,health,safety,ehs, assessment,database,branch,ohs,occupational,university,technology,sydney,australia
 hazards,first,aid,oms,online,management,system,responsibilities,legal" />
<meta name="author" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
<!-- Dublin Core Metadata -->
<meta name="dc.title" content="Online Risk Register - Safety &amp; Wellbeing, University of Technology, Sydney" />
<meta name="dc.subject" content="risk,register,assessment,environment,health,safety,ehs,branch,ohs,occupational,university,technology,sydney,human,resources,unit,hru,australia
 hazards,oms,online,management,system" />
<meta name="dc.description" content="The Online Risk Register of the University of Technology, Sydney." />
<meta name="dc.date" scheme="ISO8601" content="3 May 2006 11:13 AM" />
<meta name="dc.identifier" scheme="URL" content="http://www.ehs.uts.edu.au" />
<meta name="dc.creator" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
<meta name="dc.creator.Address" content="safetyandwellbeing@uts.edu.au" />
<meta name="dc.language" scheme="RFC1766" content="en" />
<meta name="dc.publisher" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
<title>UTS: Environment, Health and Safety - Online Risk Register - Search Header</title>
<script type="text/javascript">
function goBack()
{
  history.back(-1); 
}
</script>
<%
dim val 
dim val2
dim strTask
dim numBuildingId
dim numCampusId 
dim strSupervisor
dim numFacultyId
dim numFacilityId
dim flg
dim f
dim fc
dim s
dim c
dim b
dim conn
dim intSearchType


set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

flg = false

strTask = Session("hdnHtask")
numBuildingId =  Session("hdnBuildingId")
numCampusId = Session("hdnCampusId")
numFacultyId = Session("hdnFacultyId")
numFacilityId =Session("hdnFacilityId")
strSupervisor =Session("hdnSupervisor")

intSearchType = Session("intSearchType")

'response.write(strTask)%>
<%
'response.write(numBuildingId)%>
<%
'response.write(numFacultyId)%>
<%
'response.write(numFacilityId)%>
<%
'response.write(numCampusId)%>
<%
'response.write(strSupervisor)%>
<%

%>
<base target="main" />
<link rel="stylesheet" type="text/css" href="orr.css" media="all" />
</head>
<body>
<div id="wrapper">
  <div id="content">
    <%
	' Changes the heading to be either Action Status Report or Search depending on button pressed
	if intSearchType = 1 then
		%>
    <h1 class="pagetitle">Action Status Report</h1>
    <%
	else
		%>
    <h1 class="pagetitle">Risk Assessment Search Results</h1>
    <%	
	end if
%>
    <center>
      <table>
        <tr>
			<td>
				<form method="post" action="--WEBBOT-SELF--">
				<!--webbot bot="SaveResults" U-File="fpweb:///_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" -->
				<input type="button" value="Back to Search" name="btnBack" onclick="goBack()" />
				</form>
			</td>
			<td>
				<form method="post" target="_blank" action="SearchResultsFromMenuPrint.asp">
				<input type="submit" value="Print Preview" name="btnPrintPreview" />
				</form>
			</td>
        </tr>
      </table>
    </center>
  </div>
  <!-- close #content -->
</div>
<!-- close #wrapper -->
</body>
</html>
