<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Risk Assessment Added</title>
    <!--#include file="bootstrap.inc"--> 
</head>
<%

'strCampusname = request.form("hdnNumCampusname")
'strBuildingname = request.form("hdnNumBuildingname")
'strFacilityname = request.form("hdnNumFacilityname")

'***********************Database variables declaration**********************************************************
Dim numFacilityId
dim numCampusId
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
Dim dtDateCreated
Dim dtDate
dim numNavigationCnt
dim strConsequence
dim boolSWMSRequired

numNavigationCnt = 1
'******************** code to fetch the values from the Create QORA Form **************************************
numBuildingId = request.form("hdnBuildingId")
numFacilityId = request.form("hdnFacilityId")
numFacultyId = request.form("hdnFacultyId")
strLoginId = Request.Form("hdnLoginId")

dim operationId
dim searchType
numOperationID = request.form("operationId")

searchType = request.form("searchType")

if(searchType = "location") then
	strSQL ="Select * from tblBuilding where numBuildingId ="&numBuildingId
	
	  set connC = Server.CreateObject("ADODB.Connection")
	  connC.open constr
	  
	set rsCampus = Server.CreateObject("ADODB.Recordset")
	rsCampus.Open strSQL, connC, 3, 3
	
	numCampusId = rsCampus("numCampusId")
	'Response.Write(numCampusId) 
	
	numOperationID = 0
end if

if(searchType = "operation") then
	numFacilityID = 0
	numCampusId = 0
	numBuildingId = 0
end if

boolSWMSRequired = Request.form("boolSWMSRequired")
dtDateCreated = Request.form("txtDateCreated") 

strAssessor = Request.form("txtAssessor")
temp = instr(1,strAssessor,"'",vbTextCompare)
      if temp <> 0 then 
         strAssessor = Replace(strAssessor,"'","''",1)
      end if

strtaskDesc = Request.form("txttaskDesc")
temp = instr(1,strtaskDesc,"'",vbTextCompare)
      if temp <> 0 then 
         strtaskDesc = Replace(strtaskDesc,"'","''",1)
      end if
      
strConsultation = request.form("strConsultation")
temp = instr(1,strConsultation,"'",vbTextCompare)
      if temp > 0 then 
         strConsultation = Replace(strConsultation,"'","''",1)
      end if

//Hazards
strT1 = Request.form("T1")
temp = instr(1,strT1,"'",vbTextCompare)
      if temp <> 0 then 
         strT1 = Replace(strT1,"'","''",1)
      end if

strT2 = Request.form("T2")
temp = instr(1,strT2,"'",vbTextCompare)
      if temp <> 0 then 
         strT2 = Replace(strT2,"'","''",1)
      end if
     
//Inherent Risks 
strT3 = Request.form("T3")
temp = instr(1,strT3,"'",vbTextCompare)
      if temp <> 0 then 
         strT3 = Replace(strT3,"'","''",1)
      end if

'boolRisk = Request.form("radios")
'select case boolRisk
'  case "First" : strRisk = "H"
'  case "Second" : strRisk = "M"
'  case "Third" : strRisk = "L"
'end select 

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
'response.write(temp)
      if temp <> 0 then 
         strGenCom = Replace(strGenCom,"'","''",1)
      end if


dtDate = Request.form("txtDtActionsCompleted")

'This is our risk matrix. 
' It looks like this:
' 	1		2		3			4			5
'5 High	   High	   Extreme	 Extreme	Extreme
'4 Medium  High    High	 	 Extreme	Extreme
'3 Low	   Medium	High	 Extreme	Extreme
'2 Low	   Low		Medium	 High		Extreme
'1 Low	   Low		Medium	 High		High

'This code grabs the value form the form prior, then maps it to a
' value that our matrix can use.
likelyhood = Request.form("radiol")
consequence = Request.form("radioc")
dim likelyhoodnum
dim consequencenum

select case likelyhood
  case "Rare" 			: likelyhoodnum = 1
  case "Unlikely" 		: likelyhoodnum = 2
  case "Possible" 		: likelyhoodnum = 3
  case "Likely" 		: likelyhoodnum = 4
  case "Almost Certain"	: likelyhoodnum = 5
end select 

select case consequence
  case "Insignificant"	: consequencenum = 1
  case "Minor" 			: consequencenum = 2
  case "Moderate" 		: consequencenum = 3
  case "Major" 			: consequencenum = 4
  case "Catastrophic"	: consequencenum = 5
end select 

'We stil retain the [LMHE] code for reporting/sorting.
Dim Matrix(5, 5)
Matrix(1, 1) = "L" 
Matrix(1, 2) = "L" 
Matrix(1, 3) = "L" 
Matrix(1, 4) = "M" 
Matrix(1, 5) = "M"
 
Matrix(2, 1) = "L" 
Matrix(2, 2) = "L" 
Matrix(2, 3) = "M" 
Matrix(2, 4) = "M"
Matrix(2, 5) = "H" 

Matrix(3, 1) = "L"
Matrix(3, 2) = "M"
Matrix(3, 3) = "M"
Matrix(3, 4) = "H"
Matrix(3, 5) = "H"

Matrix(4, 1) = "M"
Matrix(4, 2) = "M"
Matrix(4, 3) = "H"
Matrix(4, 4) = "H"
Matrix(4, 5) = "E"

Matrix(5, 1) = "M"
Matrix(5, 2) = "H"
Matrix(5, 3) = "H"
Matrix(5, 4) = "E"
Matrix(5, 5) = "E"


strRisk = Matrix(consequencenum,likelyhoodnum)
Response.Write(strRisk)


'Here we anticipate what the next primary key is.  We cannot rely on the autoincrement as
'unfortunately MSAccess doesn't support @IDENTITY or any other flags to be able to recover the new
'primary key to use in the child record.
  set conn2 = Server.CreateObject("ADODB.Connection")
  conn2.open constr
  set rsPrikey = Server.CreateObject("ADODB.Recordset")
  strSQL2 ="Select max(numQORAID)+1 as numQORAID from tblQORA"
  'Response.Write strSQL2
  rsPrikey.Open strSQL2, conn2, 3, 3
  strPrikey = rsPrikey("numQORAID")

'For i = 1 To Request.Form("controls").Count 
  '  Response.Write Request.Form("controls")(i) & "<BR>" 
  
'    Next
'if dtDate = " " then
' dtDate = " "
'end if
'*************************Database connectivity Code***********************************************************

Dim conn
Dim rsAdd
Dim conn2
Dim rsAddControls

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  'conn.open constr
 
  ' setting up the recordset
'***************************Insert into database**************************************************************
   
      strSQL ="Insert into tblQORA(numQORAID, numFacilityId,dtDateCreated,strAssessor,strTaskDescription, "_
   &" strAssessRisk, strConsequence, strLikelyhood, strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
   &" boolFurtherActionsGeneralRA,strText,numFacultyId,strSupervisor,strDateActionsCompleted, strConsultation, boolSWMSRequired, strInherentRisk, numOperationId) Values "_
   &" ("& strPrikey  &","_
   &" "& numFacilityId  &","_
   &" '"& dtDateCreated &"',"_
   &" '"& strAssessor &"',"_
   &" '"& strtaskDesc &"',"_
   &" '"& strRisk &"',"_
   &" '"& consequence &"',"_
   &" '"& likelyhood &"',"_
   &" '"& strT2 &"',"_
   &" '"& strT1 &"',"_
   &" "& boolSwms &","_
   &" "& boolCRA &","_
   &" "& boolGRA &","_
   &" '"& strGenCom &"',"_
   &" "& numFacultyId &","_
   &" '"& strLoginId &"',"_
   &" '"& dtDate &"' ,"_
   &" '"& strConsultation &"' ,"_
   &" "& boolSWMSRequired &" ,"_
   &" '"& strT3 &"' ,"_
   &" "&numOperationID&" ) "
   
   set rsAdd = Server.CreateObject("ADODB.Recordset")
  'Response.Write(strSQL)
  
  'Response.End
  
  conn.open constr
  conn.BeginTrans
  conn.Execute strSQL
  conn.commitTrans
  
  'set rsCommit = Server.CreateObject("ADODB.Recordset")
  'strSQL = "COMMIT TRANSACTION"
  'rsCommit.Open strSQL, conn, 3, 3
    
  'Setup to add the Risk controls
  'We want the prikey of the record created above, which will be the most recent record created by this login.
  strHeader ="Insert into tblRiskControls(numQORAID,strControlMeasures,boolImplemented) Values ("
  strSQL2 = ""
  
  set rsAddControls = Server.CreateObject("ADODB.Connection")
  rsAddControls.open constr
  i=3
  '  while Request.Form("txtRow"&i).Count <> 0
   ' This is a quick shortcut.  There is the case where a user deletes a record 
   ' that is not the last in the list, and hence we need to know the size of the table
   ' from the previous form.
   ' I cannot find any way to do this easily, so I boldly estimate that no QORA
   ' will have more than 40 risk controls added to it, ergo the below.
  'while Request.Form("txtRow"&i).Count <> 0
  withDate ="Insert into tblRiskControls(numQORAID,strControlMeasures,boolImplemented, dtProposed) Values ("
  'Special header in the case we have no date to insert
  withoutDate = "Insert into tblRiskControls(numQORAID,strControlMeasures,boolImplemented) Values ("
  strSQL2 = ""
  
  set rsAddControls = Server.CreateObject("ADODB.Connection")
  rsAddControls.open constr
    i=3
'  while Request.Form("txtRow"&i).Count <> 0

   while i < 40
   	 if Request.Form("txtRow"&i).Count <> 0 then
   	 
  		strControl = Request.Form("txtRow"&i)(1) 
  		temp = instr(1,strControl,"'",vbTextCompare)
      	if temp <> 0 then 
         strControl = Replace(strControl,"'","''",1)
      	end if
      	
      	implemented = false
      	if Request.Form("selRow"&i).Count <> 0 then
      		implemented = true
        end if
        
        dtProposed = ""
        if Request.Form("dateRow"&i).Count <> 0 then
       		dtProposed = Request.Form("dateRow"&i)(1)
        end if
        
        if dtProposed = "" then
      		strSQL2 = withoutDate& strPrikey & " , '"&strControl &"' , "&implemented &" );"
      		
      	else
      		strSQL2 = withDate& strPrikey & " , '"&strControl &"' , "&implemented &" , '"& dtProposed &"' );"
      		
      	end if
		'response.write(strSQL2)
		rsAddControls.execute(strSQL2) 
	  end if
	  i=i+1
  wend
  rsAddControls.close 
  set rsAddControls=nothing 
  'Write the new records to DB
  
  'Response.Write(strSQL2)
  'rsAddControls.Open strSQL2, conn, 3, 3

%>

<!-- this file when imported, sets the review date in the RA -->
<% dim testval
	testval = strPrikey
%>
<!--#INCLUDE FILE="UpdateReview.asp"--> 

<body>
     <!--#include file="HeaderMenu.asp"--> 
<div id="wrapper">
  <div id="content">
    <h2 class="pagetitle">Risk Assessment <%=strPrikey%> - "<%response.write(strtaskDesc)%>" has been added successfully</h2>
	
    <div class="addAnother">

        <!--<form method="POST" action="locationSup.asp">
              <input type="hidden" name = cboBuilding value ="<%=numBuildingID%>">
              <input type="hidden" name = cboCampus value ="<%=numCampusID%>">
              <input type="hidden" name = cboRoom value ="<%=numFacilityID%>">
              <input type="submit" class="btn btn-primary" href="LocationSup.asp" value="Next" name="btnAddMore">
            </form>
        -->
         <%
            dim action
            if Session("mostRecentSearch") <> "" then   
                action = Session("mostRecentSearch")
            else
                action = "Home.asp"
            end if
            %>

          <form id="refreshResults" action="<%=action %>" method="post">
            <input type="hidden" name="confirmationMsg" value="" />
            <input type="hidden" name="searchType" value="<%=session("searchType") %>" />
            <input type="hidden" name="cboOperation" value="<%=session("cboOperation")  %>" />
            <input type="hidden" name="cboFacility" value="<%=session("cboFacility") %>" />
            <input type="hidden" name="hdnFacultyId" value="<%=session("cboFaculty") %>" />
            <input type="hidden" name="hdnBuildingId" value="<%=session("hdnBuildingId") %>" />
            <input type="hidden" name="hdnFacilityId" value="<%=session("hdnFacilityId") %>" />
            <input type="hidden" name="hdnCampusId" value="<%=session("hdnCampusId") %>" />
             <input type="hidden" name="txtHazardoustask" value="<%=session("hdnHTask") %>" />
             <input type="hidden" name="cboSupervisorName" value="<%=session("cboSupervisorName") %>" />
             <input type="submit" class="btn btn-primary" value="Next" name="btnAddMore">
        </form>
    </div>

  </div>
</div>
</body>

<!-- previously used to include file="reportAfterEdit.asp" here to show results list-->


</body>
</html>

