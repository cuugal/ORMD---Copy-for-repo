<!--#include file="aspJSON.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->
<%



	dim mode, strSuperv
	mode = request("mode")
    qora = request("qora")
    super = request("superv")


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

	if mode = "archive" then

        strSQL = "Select numOperationId from tblOperations where numFacilitySupervisorId ="& cint(super)&""_
            &" and strOperationName contains 'Archive%'"
        set rsOps = Server.CreateObject("ADODB.Recordset")
        rsOps.Open strSQL, con, 3, 3

        dim archiveId
        archiveId = 0
        while not rsOps.Eof
            archiveId = cint(rsOps("numOperationId"))
            rsOps.Movenext	
		wend 

        'if the operation ID can't be found, make a new one
        if archiveId = 0 then
            set conn = Server.CreateObject("ADODB.Connection")
            dim name
            name = "Archive - "&session("strName")
            strSql = "insert into tblOperations(numFacilitySupervisorId, strOperationName) values ("_
                    &cint(super)&" , '"&name&"')"
		    conn.open constr
		    conn.BeginTrans
		    conn.Execute strSQL
		    conn.commitTrans
            Set rs1 = conn.Execute("Select @@Identity")
            archiveId = rs1.Fields(0).Value 'new PK
        end if

        ' update the RA to have the operation ID as the operation ID

        strSQL = "update tblQORA set numFacilityId = 0, numOperationId = "&cint(archiveId)&" where numQORAId = "&cint(qora)
        set conn = Server.CreateObject("ADODB.Connection")
        conn.open constr
		conn.BeginTrans
		conn.Execute strSQL
		conn.commitTrans


        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", qora

        End With
		Response.Write oJSON.JSONoutput()  

	end if
	
	

%>