<!--#include virtual="/aspJSON.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->
<%

	dim mode, strSuperv
	mode = request("mode")
    strSuperv = request("strSuperv")
    numFacultyId = request("numFacultyId")

	set con	= server.createobject ("adodb.connection")
     con.open constr
		
	
		  
		  
	if mode = "MenuOperation" then
        strSQL = "Select numOperationID, strOperationName , strGivenName, strSurname from tblOperations, tblFacilitySupervisor, tblFaculty where tblFaculty.numFacultyID="& numFacultyID&" and tblFacilitySupervisor.numSupervisorId = tblOperations.numFacilitySupervisorId and tblFaculty.numFacultyId = tblFacilitySupervisor.numFacultyID order by strOperationName"     
        set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3
        Dim operation_name, numopId
        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", oJSON.Collection()

            dim counter
            counter = 0
            With oJSON.data("result")
                while not rsFillOperation.Eof
                    if len(strSuperv) <= 1 then
                        operation_name =rsFillOperation("strOperationName")&" - "&rsFillOperation("strGivenName")&" "&rsFillOperation("strSurname") 
                    else
                        operation_name =rsFillOperation("strOperationName")
                    end if	

                    if isNull(rsFillOperation("numOperationID")) then
                        numopid = 0
                    else
                        numopid = rsFillOperation("numOperationID")
                    end if
                    .Add Cstr(numopid), operation_name
                    '.Add counter, oJSON.Collection()                  'Create object
                   ' With .item(counter)
        
                   '     .Add "opid", numopid
                   '     .Add "operation_name", operation_name
                   ' End With
                    'response.write( "**"&numopid&"**"&operation_name&"--\n")
                    
                    counter = counter + 1
                    rsFillOperation.Movenext
                wend
            end with

        End With
		Response.Write oJSON.JSONoutput()  
	end if
	
	
	
%>