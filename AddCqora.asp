<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Online Risk Register</title>
</head>
<%

'strCampusname = request.form("hdnNumCampusname")
'strBuildingname = request.form("hdnNumBuildingname")
'strFacilityname = request.form("hdnNumFacilityname")

'***********************Database variables declaration**********************************************************
Dim numFacilityId
Dim strLoginId
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
dim numNavigationCnt
numNavigationCnt = 1

'******************** code to fetch the values from the Create QORA Form **************************************
numCampusId = request.form("hdnCampusId")
numBuildingId = request.form("hdnBuildingId")
numFacilityId = request.form("hdnFacilityId")
numFacultyId = request.form("hdnFacultyId")
strLoginId = Request.Form("hdnLoginId")

'Response.Write(strLoginId)
%>
<%

dtDateCreated = date

strAssessor = Request.form("txtAssessor")
temp = instr(1,strAssessor,"'",vbTextCompare)
      if temp > 1 then 
         strAssessor = Replace(strAssessor,"'","''",1)
      end if

strtaskDesc = Request.form("txttaskDesc")
temp = instr(1,strtaskDesc,"'",vbTextCompare)
      if temp > 1 then 
         strtaskDesc = Replace(strtaskDesc,"'","''",1)
      end if

strT1 = Request.form("T1")
temp = instr(1,strT1,"'",vbTextCompare)
      if temp > 1 then 
         strT1 = Replace(strT1,"'","''",1)
      end if

strT2 = Request.form("T2")
temp = instr(1,strT2,"'",vbTextCompare)
      if temp > 1 then 
         strT2 = Replace(strT2,"'","''",1)
      end if

boolRisk = Request.form("radios")
select case boolRisk
  case "First" : strRisk = "H"
  case "Second" : strRisk = "M"
  case "Third" : strRisk = "L"

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
temp = instr(1,strGenCom,"'",vbTextCompare)
      if temp > 1 then 
         strGenCom = Replace(strGenCom,"'","''",1)
      end if
dtDate = Request.Form("txtDtActionsCompleted") 

'*************************Database connectivity Code***********************************************************

Dim conn
Dim rsAdd

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
  ' setting up the dset
'***************************Insert into database**************************************************************
 
   strSQL ="Insert into tblQORA(numFacilityId,numCampusID,numBuildingId,dtDateCreated,strAssessor,strTaskDescription, "_
   &" strAssessRisk,strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
   &" boolFurtherActionsGeneralRA,strText,numFacultyId,strSupervisor,strDateActionsCompleted) Values "_
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
   &" "& numFacultyId &","_
   &" '"& strLoginId &"',"_
   &" '"& dtDate &"' )"

   set rsAdd = Server.CreateObject("ADODB.Recordset")
   'response.write(strSQL)
   rsAdd.Open strSQL, conn, 3, 3
   
%>
<body>
<div id="wrapper">
  <div id="content">
    <h2 class="pagetitle">Risk Assessment added successfully!</h2>
    <strong>Do you want to add another Risk Assessment? If yes, please click the 'Add more' button below.</strong>
    <center>
      <table class="bluebox" style="width: 30%; " id="table1" >
        <tr>
          <td style="text-align: center; vertical-align: middle;"><br />
            <form method="post" action="LocationAdmin.asp">
              <input type="hidden" name="hdnBuildingId" value="<%=numBuildingID%>" />
              <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" />
              <input type="hidden" name="cboRoom" value="<%=numFacilityID%>" />
              <input type="submit" value="Add more" name="btnAddMore" />
            </form></td>
          <td style="text-align: center; vertical-align: middle;"><br />
            <form method="post" action="MyQoraAdmin.asp">
              <input type="hidden" name="hdnFacilityId" value ="<%=numFacilityID%>">
              <input type="hidden" name="hdnTaskDesc" value ="<%=strTaskDesc%>">
              <input type="hidden" name="hdnNavigationCnt" value ="<%=numNavigationCnt%>">
              <input type="submit" value="Cancel" name="btnClear" />
            </form></td>
        </tr>
      </table>
    </center>
  </div>
</div>
</body>
</html>
