<!--#include file="aspJSON.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->
<%

	dim mode, strSuperv
	mode = request("mode")
    strSuperv = request("strSuperv")
    numFacultyId = request("numFacultyId")
    numBuildingId = request("numBuildingId")

	set con	= server.createobject ("adodb.connection")
     con.open constr
		  
	if mode = "Operation" then
        strSQL = "Select numOperationID, strOperationName , strGivenName, strSurname from tblOperations, tblFacilitySupervisor, "_
                &"tblFaculty where tblFaculty.numFacultyID="& numFacultyID&" and tblFacilitySupervisor.numSupervisorId = tblOperations.numFacilitySupervisorId "_
                &"and tblFaculty.numFacultyId = tblFacilitySupervisor.numFacultyID order by strOperationName"     
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
                    if len(strSuperv) >= 1 then
                        operation_name =rsFillOperation("strOperationName")&" - "&rsFillOperation("strGivenName")&" "&rsFillOperation("strSurname") 
                    else
                        operation_name =rsFillOperation("strOperationName")
                    end if	

                    if isNull(rsFillOperation("numOperationID")) then
                        numopid = 0
                    else
                        numopid = rsFillOperation("numOperationID")
                    end if

                    .Add counter, oJSON.Collection()
                    With .item(counter)
                        .Add Cstr(numopid), operation_name
                     end with              

                    counter = counter + 1
                    rsFillOperation.Movenext
                wend
            end with

        End With
		Response.Write oJSON.JSONoutput()  
	end if
	
			  
	if mode = "Supervisor" then
        strSQL ="Select * from tblFacilitySupervisor where numFacultyId ="&numFacultyId &" and boolDeprecated = 0 order by strGivenName "
        set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3
        Dim super_name
        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", oJSON.Collection()

            
            counter = 0
            With oJSON.data("result")
                while not rsFillOperation.Eof
                    
                    super_name =rsFillOperation("strGivenName")&" "&rsFillOperation("strSurname") 
                    if isNull(rsFillOperation("strLoginId")) then
                        numopid = 0
                    else
                        numopid = rsFillOperation("strLoginId")
                    end if
                    
                    .Add counter, oJSON.Collection()
                    With .item(counter)
                    .Add Cstr(numopid), super_name
                    end with

                    counter = counter + 1
                    rsFillOperation.Movenext
                wend
            end with

        End With
		Response.Write oJSON.JSONoutput()  
	end if

    if mode = "LocationBuilding" then
       strSQL = "Select  distinct(tblBuilding.strBuildingName) as strBuildingName, tblFacility.numBuildingId as NumBuildingID, tblCampus.strCampusName "_
                                    &"from tblBuilding,tblCampus,tblFacility, tblFacilitySupervisor "_
                                    &"where tblFacilitySupervisor.numFacultyId="& numFacultyID&" "_
                                    &"and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID "_
                                    &"and tblFacility.numBuildingId = tblBuilding.numBuildingId "_
                                    &"and tblBuilding.numCampusId = tblCampus.numCampusId "_
                                    &" order by strBuildingName "

   
        set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3
        Dim building_name
        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", oJSON.Collection()

            
            counter = 0
            With oJSON.data("result")
                while not rsFillOperation.Eof
                    
                    building_name = cstr(rsFillOperation("strBuildingName")) + " - " + cstr(rsFillOperation("strCampusName")) 
                    if isNull(rsFillOperation("numBuildingID")) then
                        numopid = 0
                    else
                        numopid = rsFillOperation("numBuildingID")
                    end if
                    .Add counter, oJSON.Collection()
                        With .item(counter)
                        .Add Cstr(numopid), building_name
                    end with

                    counter = counter + 1
                    rsFillOperation.Movenext
                wend
            end with

        End With
		Response.Write oJSON.JSONoutput()  
	end if


     if mode = "LocationRoom" then
        strSQL ="SELECT tblFacility.strRoomNumber,tblFacility.strRoomName,"_
                                    &" tblBuilding.strBuildingName,tblFacility.numFacilityId, strGivenName, strSurname"_
                                    &" FROM tblFacility, tblBuilding, tblFacilitySupervisor , tblFaculty"_ 
                                    &" WHERE tblFacility.numBuildingID=tblBuilding.numBuildingID "_
                                    &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
                                    &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_
                                    &" and tblFaculty.numFacultyID = "& numFacultyID&" "_
                                    &" And  tblBuilding.numBuildingId = "& numBuildingId &" "_
                                    &" order by tblFacility.strRoomNumber"

        set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3
        Dim facility_name
        Set oJSON = New aspJSON
        With oJSON.data
            .Add "result", oJSON.Collection()

            
            counter = 0
            With oJSON.data("result")
                while not rsFillOperation.Eof
                    
                    if len(strSuperv) <= 1 then
                        facility_name =cstr(rsFillOperation("strRoomNumber"))+ " - "+cstr(rsFillOperation("strRoomName"))&" - "&rsFillOperation("strGivenName")&" "&rsFillOperation("strSurname")
                    else
                        facility_name =cstr(rsFillOperation("strRoomNumber"))+ " - "+cstr(rsFillOperation("strRoomName"))
                    end if	

                    if isNull(rsFillOperation("numFacilityId")) then
                        numopid = 0
                    else
                        numopid = rsFillOperation("numFacilityId")
                    end if

                    .Add counter, oJSON.Collection()
                    With .item(counter)
                        .Add Cstr(numopid), facility_name
                     end with

                    counter = counter + 1
                    rsFillOperation.Movenext
                wend
            end with

        End With
		Response.Write oJSON.JSONoutput()  
	end if
	
%>