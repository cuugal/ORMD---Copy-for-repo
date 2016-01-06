<%
'If we have been redirected from a search form, there is no need to save QORA information prior.  Skip all this if so.
if NoSaveBeforeSWMS <> "nosave" then

	if testval ="" then
			
		'Here we anticipate what the next primary key is.  We cannot rely on the autoincrement as
		'unfortunately MSAccess doesn't support @IDENTITY or any other flags to be able to recover the new
		'primary key to use in the child record.
		  set conn2 = Server.CreateObject("ADODB.Connection")
		  conn2.open constr
		  set rsPrikey = Server.CreateObject("ADODB.Recordset")
		  strSQL2 ="Select max(numQORAID)+1 as numQORAID from tblQORA"
		  'Response.Write strSQL2
		  rsPrikey.Open strSQL2, conn2, 3, 3
		  testval = rsPrikey("numQORAID")
	end if
	
	strSQL ="Select numQORAID from tblQORA where numQORAID = "&testval
	set check_exists = server.CreateObject("ADODB.Connection")
	check_exists.Open constr
	set check_exists_results = server.CreateObject("ADODB.Recordset")
	check_exists_results.Open strSQL, check_exists, 3, 3
	
		Dim Matrix(5, 5)
		Matrix(1, 1) = "L" 
		Matrix(1, 2) = "L" 
		Matrix(1, 3) = "L" 
		Matrix(1, 4) = "M" 
		Matrix(1, 5) = "H"
		 
		Matrix(2, 1) = "L" 
		Matrix(2, 2) = "L" 
		Matrix(2, 3) = "M" 
		Matrix(2, 4) = "H" 
		Matrix(2, 5) = "H" 
		
		Matrix(3, 1) = "M" 
		Matrix(3, 2) = "M" 
		Matrix(3, 3) = "H" 
		Matrix(3, 4) = "H" 
		Matrix(3, 5) = "E" 
		
		Matrix(4, 1) = "H" 
		Matrix(4, 2) = "H" 
		Matrix(4, 3) = "E" 
		Matrix(4, 4) = "E" 
		Matrix(4, 5) = "E" 
		
		Matrix(5, 1) = "H" 
		Matrix(5, 2) = "E" 
		Matrix(5, 3) = "E" 
		Matrix(5, 4) = "E" 
		Matrix(5, 5) = "E" 
	
	'if we find a record, then we need to update it
	if not check_exists_results.EOF then
	
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
		
		'get the values from the form, then marry this up to the risk value
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
		      
		'strDate = request.Form("txtDtActionsCompleted")
		' &"strDateActionsCompleted = '"& strdate &"',"_ 
		'strDate = CDate(strDate)
		dtDateCreated = request.form("hdnDateCreated")
		dtdate = Date()
		set dcnDb = server.CreateObject("ADODB.Connection")
		dcnDb.Open constr
	
	'*************************SQL to update the database******************************************************
	
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
	                      &"strConsultation = '"& strConsultation &"',"_   
	                      &"boolSWMSRequired = "& boolSWMSRequired &","_  
	                      &"strInherentRisk = '"& strInherentRisk &"'"_                     
	                     &" Where numQORAId = "&testval  
	
	                    set rsAdd = Server.CreateObject("ADODB.Recordset")
	                    'response.write(strSQL)
	                    'rsAdd.Open strSQL, dcnDb, 3, 3 
	                   
                        rsAdd.Open strSQL, dcnDb, 3, 3 
                        dcnDb.BeginTrans
  						dcnDb.Execute strSQL
  						dcnDb.commitTrans
	                   
	                    
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
	  
	  end if
	
	'if we cannot find any record to update, then we need to create    
	if check_exists_results.EOF then
		
		numNavigationCnt = 1
		'******************** code to fetch the values from the Create QORA Form **************************************
		searchType = Session("searchType")
		numFacultyId = request.form("hdnFacultyId")
		strLoginId = Request.Form("hdnLoginId")
		
	if(searchType = "location") then
		numBuildingId = request.form("hdnBuildingId")
		numFacilityId = request.form("hdnFacilityId")
		
		
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
		numBuildingId = 0
		numFacilityId = 0
		numCampusId = 0
		numOperationID = request.form("operationId")
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
		'dim likelyhoodnum
		'dim consequencenum
		
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
		
		strRisk = Matrix(consequencenum,likelyhoodnum)
		'Response.Write(strRisk)
	
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
		   
		      strSQL ="Insert into tblQORA(numQORAID, numFacilityId,strAssessor,strTaskDescription, "_
		   &" strAssessRisk, strConsequence, strLikelyhood, strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
		   &" boolFurtherActionsGeneralRA,strText,numFacultyId,strSupervisor,dtDateCreated, strConsultation, boolSWMSRequired, strInherentRisk, numOperationID) Values "_
		   &" ("& testval  &","_
		   &" "& numFacilityId  &","_
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
		   &" '"& Date() &"' ,"_
		   &" '"& strConsultation &"' ,"_
		   &" "& boolSWMSRequired &" ,"_
		   &" '"& strT3 &"',"_
		   &numOperationID&" ) "
		   
		   set rsAdd = Server.CreateObject("ADODB.Recordset")
		  'Response.Write(strSQL)
		  'Response.end
		  conn.open constr
		  conn.BeginTrans
		  conn.Execute strSQL
		  conn.commitTrans
		    
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
	
		end if  
		
	end if
	  

	    	%>
   	
<!--The last thing to do in either case:  update the review date -->
 
<!--#INCLUDE FILE="UpdateReview.asp"--> 