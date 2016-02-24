<!--#include file="aspJSON.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->
<%



	dim mode, strSuperv
	mode = request("mode")
    cboFacility = request("cboFacility")
    cboOperation = request("cboOperation")
    qora = request("qora")


	set con	= server.createobject ("adodb.connection")
    con.open constr
    
    Dim conn
	Dim rsAdd
	Dim conn2
	Dim rsAddControls
	Dim newQORAId
    dim riskControls
    dim rows
    dim dte

	if mode = "operation" then


        strSQL = "Select * from tblQora "_
                &" where numQoraId="& qora   
        set rsFill = Server.CreateObject("ADODB.Recordset")
        rsFill.Open strSQL, con, 3, 3

		
		'Database Connectivity Code 
		  set conn = Server.CreateObject("ADODB.Connection")
		  'conn.open constr
		 
		  ' setting up the recordset
		'***************************Insert into database**************************************************************
		   
		   strSQL ="Insert into tblQORA(numFacilityId,dtDateCreated,strAssessor,strTaskDescription, "_
		   &" strAssessRisk, strConsequence, strLikelyhood, strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
		   &" boolFurtherActionsGeneralRA,strText,numFacultyId,strSupervisor,strDateActionsCompleted, strConsultation, boolSWMSRequired, "_
           &" strInherentRisk, dtReview, strJobSteps, numOperationID) Values "_
		   &" (0,"_
		   &" '"& rsFill("dtDateCreated") &"',"_
		   &" '"& rsFill("strAssessor") &"',"_
		   &" '"& rsFill("strTaskDescription") &"',"_

		   &" '"& rsFill("strAssessRisk") &"',"_
		   &" '"& rsFill("strConsequence") &"',"_
		   &" '"& rsFill("strLikelyhood") &"',"_
		   &" '"& rsFill("strControlRiskDesc") &"',"_
		   &" '"& rsFill("strHazardsDesc") &"',"_
		   &" "& rsFill("boolFurtherActionsSWMS") &","_
		   &" "& rsFill("boolFurtherActionsChemicalRA") &","_

		   &" "& rsFill("boolFurtherActionsGeneralRA") &","_
		   &" '"& rsFill("strText") &"',"_
		   &" "& rsFill("numFacultyId") &","_
		   &" '"& rsFill("strSupervisor") &"',"_
		   &" '"& rsFill("strDateActionsCompleted") &"' ,"_
		   &" '"& rsFill("strConsultation") &"' ,"_
		   &" "& rsFill("boolSWMSRequired") &" ,"_
		   &" '"& rsFill("strInherentRisk") &"',"_
           &" '"& rsFill("dtReview") &"',"_
           &" '"& rsFill("strJobSteps") &"',"_
		   & cint(cboOperation)&" ) "
		   
		   set rsAdd = Server.CreateObject("ADODB.Recordset")
		  'Response.Write(strSQL)
		  'Response.end
		  conn.open constr
		  conn.BeginTrans
		  conn.Execute strSQL
		  conn.commitTrans
		    
        Set rs1 = conn.Execute("Select @@Identity")

        newQoraId = rs1.Fields(0).Value 'new PK
  
        

        riskControls = "Select * from tblRiskControls where numQORAID =" &qora
        set rsRisk = Server.CreateObject("ADODB.Recordset")
        rsRisk.Open riskControls, con, 3, 3

        
        
        riskControls  ="Insert into tblRiskControls(numQORAID,strControlMeasures,boolImplemented, dtProposed) Values ("

  
        rows = ""
        while not rsRisk.Eof
            if isNUll(rsRisk("dtProposed")) then
                dte = "NULL"
            else
                dte = rsRisk("dtProposed")
            end if

            rows = newQoraId&",'"&rsRisk("strControlMeasures")&"',"&rsRisk("boolImplemented")&","&dte&")"
            rows = riskControls&rows

            conn.BeginTrans
		    conn.Execute rows
		    conn.commitTrans

		rsRisk.Movenext	
		wend 



        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", newQoraId

        End With
		Response.Write oJSON.JSONoutput()  
	end if
	
			  
	if mode = "location" then
        strSQL = "Select * from tblQora "_
                &" where numQoraId="& qora   
        set rsFill = Server.CreateObject("ADODB.Recordset")
        rsFill.Open strSQL, con, 3, 3

 
		
		'Database Connectivity Code 
		  set conn = Server.CreateObject("ADODB.Connection")
		  'conn.open constr
		 
		  ' setting up the recordset
		'***************************Insert into database**************************************************************
		   
		   strSQL ="Insert into tblQORA(numFacilityId,dtDateCreated,strAssessor,strTaskDescription, "_
		   &" strAssessRisk, strConsequence, strLikelyhood, strControlRiskDesc,strHazardsDesc,boolFurtherActionsSWMS,boolFurtherActionsChemicalRA, "_
		   &" boolFurtherActionsGeneralRA,strText,numFacultyId,strSupervisor,strDateActionsCompleted, strConsultation, boolSWMSRequired, "_
           &" strInherentRisk, dtReview, strJobSteps, numOperationID) Values "_
		   &" ("&cint(cboFacility)&","_
		   &" '"& rsFill("dtDateCreated") &"',"_
		   &" '"& rsFill("strAssessor") &"',"_
		   &" '"& rsFill("strTaskDescription") &"',"_

		   &" '"& rsFill("strAssessRisk") &"',"_
		   &" '"& rsFill("strConsequence") &"',"_
		   &" '"& rsFill("strLikelyhood") &"',"_
		   &" '"& rsFill("strControlRiskDesc") &"',"_
		   &" '"& rsFill("strHazardsDesc") &"',"_
		   &" "& rsFill("boolFurtherActionsSWMS") &","_
		   &" "& rsFill("boolFurtherActionsChemicalRA") &","_

		   &" "& rsFill("boolFurtherActionsGeneralRA") &","_
		   &" '"& rsFill("strText") &"',"_
		   &" "& rsFill("numFacultyId") &","_
		   &" '"& rsFill("strSupervisor") &"',"_
		   &" '"& rsFill("strDateActionsCompleted") &"' ,"_
		   &" '"& rsFill("strConsultation") &"' ,"_
		   &" "& rsFill("boolSWMSRequired") &" ,"_
		   &" '"& rsFill("strInherentRisk") &"',"_
           &" '"& rsFill("dtReview") &"',"_
           &" '"& rsFill("strJobSteps") &"',"_
		   &"0 ) "
		   
		   set rsAdd = Server.CreateObject("ADODB.Recordset")
		  'Response.Write(strSQL)
		  'Response.end
		  conn.open constr
		  conn.BeginTrans
		  conn.Execute strSQL
		  conn.commitTrans
		    
        Set rs1 = conn.Execute("Select @@Identity")
        newQoraId = rs1.Fields(0).Value 'new PK
  

        riskControls = "Select * from tblRiskControls where numQORAID =" &qora
        set rsRisk = Server.CreateObject("ADODB.Recordset")
        rsRisk.Open riskControls, con, 3, 3

        
        
        riskControls  ="Insert into tblRiskControls(numQORAID,strControlMeasures,boolImplemented, dtProposed) Values ("

        rows = ""
        while not rsRisk.Eof
            if isNUll(rsRisk("dtProposed")) then
                dte = "NULL"
            else
                dte = rsRisk("dtProposed")
            end if

            rows = newQoraId&",'"&rsRisk("strControlMeasures")&"',"&rsRisk("boolImplemented")&","&dte&")"
            rows = riskControls&rows

            conn.BeginTrans
		    conn.Execute rows
		    conn.commitTrans

		rsRisk.Movenext	
		wend 



        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", newQoraId

        End With
		Response.Write oJSON.JSONoutput()  
	end if

%>