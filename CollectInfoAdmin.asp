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
    dim searchType

   if Request.QueryString("searchType") <> "" then
       strTask = Request.QueryString("txtHazardousTask")
       numBuildingId = cint("0"&Request.QueryString("hdnBuildingId"))
       numCampusId = cint("0"&Request.QueryString("hdnCampusId"))
       strSupervisor = Request.QueryString("cboSupervisorName")
       numFacultyId = cint("0"&Request.QueryString("hdnFacultyId"))
       numFacilityId = cint("0"&Request.QueryString("hdnFacilityId"))
       strOperation = cint("0"&Request.QueryString("cboOperation"))
       searchType = request.QueryString("searchType")
    else
       strTask = Request.Form("txtHazardousTask")
       numBuildingId = cint("0"&Request.form("hdnBuildingId"))
       numCampusId = cint("0"&Request.form("hdnCampusId"))
       strSupervisor = Request.form("cboSupervisorName")
       numFacultyId = cint("0"&Request.form("hdnFacultyId"))
       numFacilityId = cint("0"&Request.form("hdnFacilityId"))
       strOperation = cint("0"&Request.Form("cboOperation"))
       searchType = request.form("searchType")
    end if

 
  
   Session("hdnHTask") = strTask
   Session("hdnBuildingId") = numBuildingId
   Session("hdnCampusId") = numCampusId
   Session("hdnFacultyId") = numFacultyId
   Session("hdnFacilityId") = numFacilityId
   Session("hdnSupervisor") = strSupervisor
   Session("hdnOperationID") = strOperation
   Session("cboSupervisorName") = strSupervisor

    'persist these in session so that we can easily replicate the search results
    session("cboOperation") = strOperation
	session("cboFacility") = numFacilityId
    session("cboFaculty") = numFacultyId

    Session("mostRecentSearch") = "CollectInfoAdmin.asp"
    Session("confirmationMsg") = request.form("confirmationMsg")

   session("searchType") = searchType
   
   Response.Redirect("resultsFrameFromAdmin.asp") 
  %>
<P>&nbsp;</P>

</BODY>
</HTML>
