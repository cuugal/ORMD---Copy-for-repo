<!--#INCLUDE FILE="DbConfig.asp"-->
<%
'sOutput stores the final output
'sData stores each line output

    set con	= server.createobject ("adodb.connection")
    con.open constr



    strSQL = "SELECT numQORAId,strGivenName,  tblFacilitySupervisor.numFacultyId as numFaculty, strSurname,  strRoomName, strRoomNumber, null as strOperationName, strTaskDescription, dtReview "_
 &" FROM  tblFacility, tblQORA, tblFacilitySupervisor "_
 &" Where tblQORA.numFacilityID = tblFacility.numFacilityID"_
 &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
 &" and dtReview < Date()  "_
 
 
 &" union all "_
 
 &"SELECT numQORAId,strGivenName, tblFacilitySupervisor.numFacultyId as numFaculty, strSurname,  null as strRoomName, null as strRoomNumber, strOperationName, strTaskDescription, dtReview "_
 &" FROM tblOperations , tblQORA, tblFacilitySupervisor "_
 &" where tblQORA.numOperationID = tblOperations.numOperationId "_
 &" and tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID "_
 &" and dtReview < Date() "

	' 7June2019 DLJ added numFacultyId to report

   'response.write strSQL

   'response.end
         
        set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3

'==== write the title (name of the column) ===
    sData = Chr(34) & "First Name" & Chr(34) & ","
    sData = sData & Chr(34) & "Last Name" & Chr(34)& ","
	    sData = sData & Chr(34) & "Faculty" & Chr(34)& ","
    sData = sData & Chr(34) & "Room Name" & Chr(34)& ","
    sData = sData & Chr(34) & "Room Number" & Chr(34)& ","
    sData = sData & Chr(34) & "Operation" & Chr(34)& ","
    sData = sData & Chr(34) & "RA Number" & Chr(34)& ","
    sData = sData & Chr(34) & "Task" & Chr(34)& ","
    sData = sData & Chr(34) & "Review Date" & Chr(34)

sOutPut = sOutPut & sData & vbCrLf

     while not rsFillOperation.Eof
        '===== now output 1 line of data =======
        sData = Chr(34) & rsFillOperation("strGivenName") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("strSurname") & Chr(34) & ","
		        sData = sData &Chr(34) & rsFillOperation("numFaculty") & Chr(34) & ","
        sData =sData & Chr(34) & rsFillOperation("strRoomName") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("strRoomNumber") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("strOperationName") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("numQORAId") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("strTaskDescription") & Chr(34) & ","
        sData = sData &Chr(34) & rsFillOperation("dtReview") & Chr(34) 
     
        sOutPut = sOutPut & sData & vbCrLf
     rsFillOperation.Movenext
     wend

FileName="expiredRAReport.csv" 'default file name

Response.Clear
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "filename=" & FileName & ";"
Response.Write(sOutPut)
%>