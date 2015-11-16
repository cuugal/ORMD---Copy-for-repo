<script type="text/javascript">
function goBack()
{
  history.back(-1); 
}
</script>
<%
    
'dim val 
'dim val2
'dim strTask
'dim numBuildingIdHead
'dim numCampusId 
'dim strSupervisor
'dim numFacultyId
'dim numFacilityId
'dim flg
'dim f
'dim fc
'dim s
'dim c
'dim b
'dim conn
'dim intSearchType


'set conn = Server.CreateObject("ADODB.Connection")
'conn.open constr

'flg = false

'strTask = Session("hdnHtask")
'numBuildingId =  Session("hdnBuildingId")
'numCampusId = Session("hdnCampusId")
'numFacultyId = Session("hdnFacultyId")
'numFacilityId =Session("hdnFacilityId")
'strSupervisor =Session("hdnSupervisor")

'intSearchType = Session("intSearchType")

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

