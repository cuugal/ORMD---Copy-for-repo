<%@Language = VBscript%>
<%Response.Buffer = true%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<%
  

  dim strTask
  dim numBuildingId
  dim numCampusId 
  dim strSupervisor
  dim numFacultyId
  dim numFacilityId
  dim strOperation
   
   strTask = Request.Form("txtHazardousTask")
   strOperation = Request.Form("cboOperation")
   numBuildingId = Request.form("hdnBuildingId")
   numCampusId = Request.form("hdnCampusId")
   strSupervisor = Request.form("cboSupervisorName")
   numFacultyId = Request.form("hdnFacultyId")
   numFacilityId = Request.form("cboRoom")
   
   session("searchType") = request.form("searchType")
   
   Response.AppendToLog("***"+numCampusId+"***")
   Response.AppendToLog("***"+strSupervisor+"***")
   Response.AppendToLog("***"+numFacultyId+"***")
   Response.AppendToLog("***"+numFacilityId+"***")
    
   Session("hdnHTask") = strTask
   Session("hdnBuildingId") = numBuildingId
   Session("hdnCampusId") = numCampusId
   Session("hdnFacultyId") = numFacultyId
   Session("hdnFacilityId") = numFacilityId
   Session("hdnSupervisor") = strSupervisor
   Session("hdnOperationID") = strOperation
  
  ' Determines if normal Search or Action Status Report
  if request.form("btnSearch") = "Action Status Report" then
  	session("intSearchType") = 1
	'response.write("Action Status Report")
  else
  	session("intSearchType") = 0
	'response.write("Normal Search")
  end if
      
  if((strTask = "" or strTask = "0") and numBuildingId = "0" and numCampusId = "0" and  (strSupervisor = "" or strSupervisor = "0") and numFacultyId = "0" and numFacilityId = "0" and strOperation = "0") then
  	
    'Response.Redirect("Menu.asp")
  	%>
	<script><!--
	alert("To many results, please refine your search criteria");
	location.href="Home.asp";
	// -->
	</script>
	<%
  else
  	Response.Redirect("resultsFrame.asp") 
  end if
  %>
<P>&nbsp;</P>

</BODY>
</HTML>
