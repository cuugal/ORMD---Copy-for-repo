<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>hi</title>
</head>
<body>
    <!--#include file="adminMenu.asp" -->
<%

'strCampusname = request.form("hdnNumCampusname")
'strBuildingname = request.form("hdnNumBuildingname")
'strFacilityname = request.form("hdnNumFacilityname")

'***********************Database variables declaration**********************************************************
Dim numFacilityId
Dim strAssessor
Dim strtaskDesc
Dim strT1
Dim strT2
Dim strRisk
Dim boolRisk
Dim boolSWMS
Dim boolCRA
Dim boolGRA
Dim strGenCom
Dim dtCreated
Dim dtDate

'******************** code to fetch the values from the Create QORA Form **************************************
numCampusId = request.form("hdnCampusId")
numBuildingId = request.form("hdnBuildingId")
numFacilityId = request.form("hdnFacilityId")
numFacultyId = request.form("cboFaculty")

dtDateCreated = date
strAssessor = Request.form("txtAssessor")
strtaskDesc = Request.form("txttaskDesc")
strT1 = Request.form("T1")
strT2 = Request.form("T2")
boolRisk = Request.form("radios")
select case boolRisk
  case "First" : strRisk = "H"
  case "First" : strRisk = "M"
  case "First" : strRisk = "L"

end select 

boolswms = Request.form("notify")
if boolSwms = "ON" then
   boolSwms = "Yes"
else
   boolSwms = "No"
end if

boolCRA = Request.form("notify2")
if boolCRA = "ON" then
   boolCRA = "Yes"
else
   boolCRA = "No"
end if

boolGRA = Request.form("chkGRA")
if boolGRA = "ON" then
   boolGRA = "Yes"
else
   boolGRA = "No"
end if

strGenCom = Request.form("txtGenComments")
dtDate = Request.form("txtDate")
if dtDate = " " then
 dtDate = " "
end if
'*************************Database connectivity Code***********************************************************

Dim conn
Dim rsAdd

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
  ' setting up the recordset
'***************************Insert into database**************************************************************
 
   strSQL ="Insert into tblQORA(numFacilityId,numCampusID,numBuildingId,dtDateCreated,strAssessor,strTaskDescription, "_
   &" strAssessRisk,strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
   &" boolFurtherActionsGeneralRA,strText,dtDate,numFacultyId) Values "_
   &" ("& numFacilityId  &","_
   &" "& numCampusiD &","_
   &" "& numBuildingId &","_
   &" '"& dtDateCreated &"',"_
   &" '"& strAssessor &"',"_
   &" '"& strtaskDesc &"',"_
   &" '"& strRisk &"',"_
   &" '"& strT2 &"',"_
   &" '"& strT1 &"',"_
   &" "& boolSwms &","_
   &" "& boolCRA &","_
   &" "& boolGRA &","_
   &" '"& strGenCom &"',"_
   &" '"& dtDate &"',"_
   &" "& numFacultyId &" )"

   set rsAdd = Server.CreateObject("ADODB.Recordset")
   'response.write(strSQL)
   rsAdd.Open strSQL, conn, 3, 3
   
%>
<font color="#800000"><b>Record Edited successfully ! </b> 
<p><br>
&nbsp;</p>
<p>&nbsp;</p>

</body>

</html>