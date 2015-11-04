<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<head>
<meta http-equiv="Content-Language" content="en-au">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>
<html>
<body link="#000000" vlink="#000000" alink="#000000">
<%
' Declaring the general variables
  Dim old_Campus
  Dim delete_Campus
  
  Dim delete_Building
  Dim delete_Faculty
  Dim delete_Supervisor

  Dim old_FacilityId
  Dim old_Building
  Dim old_RoomName
  Dim old_RoomNumber
  
  Dim delete_Sup_LoginId
    
  Dim hdn_Option
  
' Declaring the Database variables  
  
  Dim conn
  
  Dim RsCheckBuilding
  Dim RsCheckFaculty
  Dim RsCheckSupervisor
  Dim RsCheckFacility
  
  Dim RsDeleteCampus
  Dim RsDeleteBuilding
  Dim RsDeleteFaculty
  Dim RsDeleteSupervisor
  Dim RsDeleteFacility
  Dim strSQL
   
' Reteieving the contents from the input forms
 
  old_Campus = Request.form("cboCampusName")
  old_Faculty = Request.form("cboFacultyName")
  old_B = Request.form("cboBuildingName")
  old_Building = Request.form("hdnBuildingID")
  old_RoomNumber = Request.form("hdnRoomNumber")
  old_RoomName = Request.form("txtRoomName")

  
  delete_Faculty = Request.form("txtFacultyName")
  delete_Supervisor = Request.form("txtSupervisorName")
  
  delete_Sup_LoginID = Request.form("hdnLoginId")
     
  hdn_Option = Request.form("hdnOption")
             
' code for the database connectivity.
  
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
   
' code to write the contents from the admin files into the database system.

  select case(hdn_Option)
    
    ' case for the campus name
      Case "Campus" : ' code to delete campus from the database
  	                  ' Checking the referential integrity
   	                        strSQL ="select numBuildingId from tblBuilding where numCampusID ="& old_Campus 
 					        set rsCheckCampus = Server.CreateObject("ADODB.Recordset")
                            rsCheckCampus.Open strSQL, conn, 3, 3 

                          if rsCheckCampus.EOF = true then
                          ' deleteing existing campus
                                                 
 					        strSQL ="delete from tblCampus where numCampusID ="& old_Campus 
 					        set rsDeleteCampus = Server.CreateObject("ADODB.Recordset")
                            rsDeleteCampus.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Campus name has been deleted successfully !</b></font></p> <%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
						                            
                          else%>
                            <p><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this campus from the database please first delete associated buildings,facility and Risk Assesments.</b></font></p> <%

                          end if     
  
   ' case for the Building name
      Case "Building" : ' code to delete Building from the database
  	                  
  	                   ' Checking the referential integrity
   	                        strSQL ="select numBuildingId from tblFacility where numBuildingID ="& old_B
 					        set rsCheckF = Server.CreateObject("ADODB.Recordset")
                            rsCheckF.Open strSQL, conn, 3, 3 

                          if rsCheckF.EOF = true then
                          ' deleteing existing Building
                                                                           
 					        strSQL ="delete from tblBuilding where numBuildingID ="& old_B 
 					        set rsDeleteBuilding = Server.CreateObject("ADODB.Recordset")
                            rsDeleteBuilding.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Building name has been deleted successfully !</b></font></p> <%
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
													 
                           else%>
                            <p><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this building from the database please first delete associated facility and Risk Assesments.</b></font></p> <%

                            
                           end if
                          
  Case "Faculty" : ' code to delete Faculty from the database
  	                  ' Checking the referential integrity
   	                        strSQL ="select numFacultyId from tblFacilitySupervisor where numFacultyID ="& old_Faculty 
 					        set rsCheckFaculty = Server.CreateObject("ADODB.Recordset")
                            rsCheckFaculty.Open strSQL, conn, 3, 3 

                          if rsCheckFaculty.EOF = true then
                          ' deleteing existing Faculty
                                                 
 					        strSQL ="delete from tblFaculty where numFacultyID ="& old_Faculty 
 					        set rsDeleteFaculty = Server.CreateObject("ADODB.Recordset")
                            rsDeleteFaculty.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Faculty / Unit name has been deleted successfully !</b></font></p> <%
                        
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
												    
                          else%>
                            <p><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this Faculty / Unit from the database please first delete associated Supervisor,campus,buildings,facility and Risk Assesments.</b></font></p> <%

                          end if     
        
   Case "Facility" : ' code to delete Faculty from the database

   	                       strSQL ="select * from tblFacility where strRoomNumber = '"& old_RoomNumber &"' and strRoomName = '"& old_RoomName &"' and numBuildingID = "& old_Building &""
   	                          	                          	                          
 					       set rsCheckFacility = Server.CreateObject("ADODB.Recordset")
                           rsCheckFacility.Open strSQL, conn, 3, 3 
                           
                           'response.write(strSQL)
                           old_FacilityId = rsCheckFacility("numFacilityId")
                           'response.write(old_FacilityId)

                           if rsCheckFacility.EOF <> true then
                            'deleteing existing Facility
                                                 
 					        strSQL ="delete from tblFacility where numFacilityID ="& old_FacilityID 
 					        set rsDeleteFacility = Server.CreateObject("ADODB.Recordset")
                            rsDeleteFacility.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Facility has been deleted successfully !</b></font></p> <%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
						                            
                          else%>
                            <p><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this Facility from the database please first delete associated Risk Assessments.</b></font></p> <%

                          end if   
                          
    Case "Operation" : ' code to delete Operation from the database
						Dim operationID
	  					
	  					operationID = Request.form("hdnOperationId")
	  						  					
	                   strSQL ="select * from tblQORA where numOperationID = "&operationID
	                          	                          	                          
					   set rsCheckOperation = Server.CreateObject("ADODB.Recordset")
                       rsCheckOperation.Open strSQL, conn, 3, 3 
                   
                       if rsCheckOperation.EOF then
                        'deleting existing Facility
                                                 
 					        strSQL ="delete from tblOperations where numOperationID ="& operationID 
 					        set rsDeleteFacility = Server.CreateObject("ADODB.Recordset")
                            rsDeleteFacility.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Operation has been deleted successfully !</b></font></p> <%

					   	'Redirect to admin functions page on success
						'response.redirect("admin.asp")
						                            
                          else%>
                            <p><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this Operation from the database please first delete associated Risk Assessments.</b></font></p> <%

                          end if  
                          

 ' case for the Facility Supervisor
      Case "Supervisor" :  ' code to delete campus from the database
      					'AA jan 2010 this is all new:  need to look up the supervisor ID from the strLoginID (relationship repair)
      					dim numLoginID
      					strSQL = "select numSupervisorID from tblFacilitySupervisor where strLoginID  = '"& delete_Sup_LoginID &"'" 
      						set rsGetSupID = Server.CreateObject("ADODB.Recordset")
                            rsGetSupID.Open strSQL, conn, 3, 3 
                            if rsGetSupID.EOF <> true then 
                            	numLoginID = rsGetSupID("numSupervisorID")
                            end if
      
  	                  ' Checking the referential integrity
  	                  'AA jan 2010 changes to relationship
   	                        'strSQL ="select strFacilitySupervisor from tblFacility where strFacilitySupervisor  = '"& delete_Sup_LoginID &"'" 
 					        strSQL ="select numFacilitySupervisorID from tblFacility where numFacilitySupervisorID  = "& numLoginID &"" 
 					        set rsCheckFacSup = Server.CreateObject("ADODB.Recordset")
                            rsCheckFacSup.Open strSQL, conn, 3, 3 

                          if rsCheckFacSup.EOF = true then
                          ' deleteing existing Facility Supervisor
                                                 
 					        strSQL ="delete from tblFacilitySupervisor where strLoginId ='"& delete_Sup_LoginID &"'" 					        
 					        set rsDeleteFacSup = Server.CreateObject("ADODB.Recordset")
                            rsDeleteFacSup.Open strSQL, conn, 3, 3 %>
                             <p><font color="#660033"><b>The Facility Supervisor has been deleted successfully !</b></font></p> <%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
						                            
                          else%>
                            <P><font color="#660033"><b>This record can't been deleted as it is associated with other records in the database! <BR>
                            To permanently delete this Supervisor from the database please first delete associated campus,buildings,facility and Risk Assesments.</b></font></P>
 <%


                          end if     
  end select    
  
%>

</body>
</html>