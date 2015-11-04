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
   
   strTask = Request.Form("txtHazardousTask")
   numBuildingId = Request.form("hdnBuildingId")
   numCampusId = Request.form("hdnCampusId")
   strSupervisor = Request.form("cboSupervisorName")
   numFacultyId = Request.form("hdnFacultyId")
   numFacilityId = Request.form("cboRoom")
   strOperation = Request.Form("cboOperation")
  
   Session("hdnHTask") = strTask
   Session("hdnBuildingId") = numBuildingId
   Session("hdnCampusId") = numCampusId
   Session("hdnFacultyId") = numFacultyId
   Session("hdnFacilityId") = numFacilityId
   Session("hdnSupervisor") = strSupervisor
   Session("hdnOperationID") = strOperation

   session("searchType") = request.form("searchType")
   
   Response.Redirect("resultsFrameFromAdmin.asp") 
  %>
<P>&nbsp;</P>

</BODY>
</HTML>
