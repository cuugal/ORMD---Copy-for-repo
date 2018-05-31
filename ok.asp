<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->

<html>

<head>
<head>
<meta http-equiv="Content-Language" content="en-au">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title></title>
    <!--#include file="bootstrap.inc"--> 
</head>
     <!--#include file="HeaderMenu.asp"--> 
<%
session.LCID = 2057	'English(British) format

Function Escape(sString)

'Replace any Cr and Lf with <br>
strReturn = Replace(sString , vbCrLf, "<br>")
strReturn = Replace(strReturn , vbCr , "<br>")
strReturn = Replace(strReturn , vbLf , "<br>")           
Escape = strReturn
End Function


'*********** declaring the variables****************************
dim testVal 
dim rsQORA
dim dcnDb
dim strSQL 

dim numCampusId
dim numBuildingId
dim numFacilityId
dim numFacultyId
dim strSupervisor
dim dtDateCreated
dim dtDate
dim strAssessor
dim strTaskDescription
dim strHazardsDesc
dim strAssessRisk
dim strControlRiskDesc
dim strText 
dim strDate
dim strInherentRisk
dim strConsequence
dim boolSWMSRequired

'boolRisk = Request.form("radios")
'select case boolRisk
'  case "First" : strRisk = "H"
'  case "Second" : strRisk = "M"
'  case "Third" : strRisk = "L"
' end select 

'This is our risk matrix. 
' It looks like this:
' 	1		2		3			4			5
'5 High	   High	    Extreme	 Extreme	Extreme
'4 Medium  High     High	 Extreme	Extreme
'3 Low	   Medium	High	 Extreme	Extreme
'2 Low	   Low		Medium	 High		Extreme
'1 Low	   Low		Medium	 High		High
' NOTE: doesnt look like this now - as of 31May2018 changed

'get the values from the form, then marry this up to the risk value
likelyhood = Request.form("radiol")
consequence = Request.form("radioc")

searchType= Session("searchType")

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
boolSWMSRequired = Request.form("boolSWMSRequired")

boolswms = Request.form("notify")
if boolSwms = "on" then
   boolSwms = "Yes"
else
   boolSwms = "No"
end if

boolCRA = Request.form("notify2")
if boolCRA = "on" then
   boolCRA = "Yes"
else
   boolCRA = "No"
end if

boolGRA = Request.form("chkGRA")
if boolGRA = "on" then
   boolGRA = "Yes"
else
   boolGRA = "No"
end if
'****************Fetching the details***************************

testval = request.form("hdnQORAID")
'response.write(testval)
testval = cint(testval)
'response.write(testval)
strAssessor = request.form("txtAssessor")
temp = instr(1,strAssessor,"'",vbTextCompare)
      if temp <> 0 then 
         strAssessor = Replace(strAssessor,"'","''",1)
      end if
      
strTaskDescription = request.form("txtTaskDesc")
session("HTask") = strTaskDescription
session("pn")= 1
temp = instr(1,strTaskDescription,"'",vbTextCompare)
      if temp > 0 then 
         strTaskDescription = Replace(strTaskDescription,"'","''",1)
         session("HTask") = strTaskDescription
      end if
      
strHazardsDesc = request.form("T1")
temp = instr(1,strHazardsDesc,"'",vbTextCompare)
      if temp > 0 then 
         strHazardsDesc = Replace(strHazardsDesc,"'","''",1)
      end if
      
strConsultation = request.form("strConsultation")
temp = instr(1,strConsultation,"'",vbTextCompare)
      if temp > 0 then 
         strConsultation = Replace(strConsultation,"'","''",1)
      end if

strControlRiskDesc =request.Form("T2")
temp = instr(1,strControlRiskDesc,"'",vbTextCompare)
      if temp > 0 then 
         strControlRiskDesc = Replace(strControlRiskDesc,"'","''",1)
      end if
      
strInherentRisk =request.Form("T3")
temp = instr(1,strInherentRisk,"'",vbTextCompare)
      if temp > 0 then 
         strInherentRisk = Replace(strInherentRisk,"'","''",1)
      end if
      
strText = request.Form("txtGenComments")
temp = instr(1,strText,"'",vbTextCompare)
      if temp > 0 then 
         strText = Replace(strText,"'","''",1)
      end if
      
strDate = request.Form("txtDtActionsCompleted")
'strDate = CDate(strDate)
dtDateCreated = request.form("hdnDateCreated")
dtdate = date
set dcnDb = server.CreateObject("ADODB.Connection")
dcnDb.Open constr

'*************************SQl to update the database******************************************************

                         strSQL = "Update tblQORA Set "_
                          &"strAssessor = '"& strAssessor &"',"_
                          &"strTaskDescription = '"& strTaskDescription &"',"_
                          &"strAssessRisk = '"&strRisk &"',"_
                          &"strConsequence = '"&Consequence &"',"_
                          &"strLikelyhood = '"&Likelyhood &"',"_ 
                          &"strControlRiskDesc = '"& strControlRiskDesc &"',"_
                          &"strHazardsDesc = '"& strHazardsDesc  &"',"_
                          &"boolFurtherActionsSWMS = "& boolSWMS  &","_
                          &"boolFurtherActionsChemicalRA = "& boolCRA  &","_
                          &"boolFurtherActionsGeneralRA = "& boolGRA &","_
                          &"strText = '"& strText &"',"_
                          &"dtDateCreated = '"& dtdate &"',"_                        
                          &"strDateActionsCompleted = '"& strdate &"',"_ 
                          &"strConsultation = '"& strConsultation &"',"_   
                          &"boolSWMSRequired = "& boolSWMSRequired &","_  
                          &"strInherentRisk = '"& strInherentRisk &"'"_                     
                         &" Where numQORAId = "&testval  
                         
                         
                     
                        set rsAdd = Server.CreateObject("ADODB.Recordset")
                        'response.write(strSQL)
                        rsAdd.Open strSQL, dcnDb, 3, 3 
                        dcnDb.BeginTrans
  						dcnDb.Execute strSQL
  						dcnDb.commitTrans
                        %>

                          
                         <%


'*********************************************************************************************************
'Delete out all the old risk controls
set remover = server.CreateObject("ADODB.Connection")
remover.open constr
strOutwithOld = "Delete * from tblRiskControls where numQORAID = "&testval
remover.execute(strOutwithOld)
remover.close 
set remover=nothing 

'Setup to add the Risk controls
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
      		strSQL2 = withoutDate& testval & " , '"&strControl &"' , "&implemented &" );"
      		
      	else
      		strSQL2 = withDate& testval & " , '"&strControl &"' , "&implemented &" , '"& dtProposed &"' );"
      		
      	end if
		'response.write(strSQL2)
		rsAddControls.execute(strSQL2) 
	  end if
	  i=i+1
  wend
  rsAddControls.close 
  set rsAddControls=nothing 
  
  


%>
<!--#INCLUDE FILE="UpdateReview.asp"--> 
<body>
<div id="wrapper">
  <div id="content">
    <!--h2 class="pagetitle">Risk Assessment <%=testval%> has been edited successfully</h2-->
    <h2 class="pagetitle">Risk Assessment <%=testval%> - "<%response.write(strTaskDescription)%>" has been edited successfully</h2>

	
      <%
          dim action
          if Session("mostRecentSearch") <> "" then
            action = Session("mostRecentSearch")
          else
            action = "Home.asp"
          end if
           %>
      <form id="refreshResults" action="<%=action %>" method="post">
          <input type="hidden" name="confirmationMsg" value="Risk Assessment <%=testval%> - <%=strTaskDescription%> has been edited successfully" />
		 
        <input type="hidden" name="searchType" value="<%=session("searchType") %>" />
        <input type="hidden" name="cboOperation" value="<%=session("cboOperation")  %>" />
        <input type="hidden" name="cboFacility" value="<%=session("cboFacility") %>" />
          <input type="hidden" name="hdnFacultyId" value="<%=session("cboFaculty") %>" />
         <input type="hidden" name="hdnBuildingId" value="<%=session("hdnBuildingId") %>" />
         <input type="hidden" name="hdnFacilityId" value="<%=session("hdnFacilityId") %>" />
          <input type="hidden" name="hdnCampusId" value="<%=session("hdnCampusId") %>" />
        <input type="hidden" name="txtHazardoustask" value="<%=session("hdnHTask") %>" />
        <input type="hidden" name="cboSupervisorName" value="<%=session("cboSupervisorName") %>" />
          <input type="submit" class="btn btn-primary" value="Next" />
    </form>
  </div>
</div>
    

</body>

</html>

