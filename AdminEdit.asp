<%@Language = VBscript%>
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
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
  Dim edit_Campus
  
  Dim old_Building
  Dim edit_Building
  Dim edit_Building_Campus
  
  Dim edit_Facility_RoomNumber
  Dim edit_Facility_RoomName
  Dim edit_FacilitySupervisor
  Dim edit_Facility_Building
  
  Dim edit_Faculty
  Dim edit_DGivenName
  Dim edit_DSurname
  Dim edit_DLogin
  Dim edit_DPassword

  Dim edit_Sup_SName
  Dim edit_Sup_GName
  Dim edit_Sup_LoginId
  Dim edit_Sup_Password
  Dim edit_Sup_Faculty

    
  Dim hdn_Option
  
' Declaring the Database variables  
  
  Dim conn
  
  Dim RsCheckBuilding
  Dim RsCheckFaculty
  Dim RsCheckSupervisor
  Dim RsCheckFacility
  Dim RsEditCampus
  Dim RsEditBuilding
  Dim RsEditFaculty
  Dim RsEditSupervisor
  Dim RsEditFacility
  Dim strSQL
 
 
  hdn_Option = Request.form("hdnOption")
  
' code for the database connectivity.
  
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
   
' code to write the contents from the admin files into the database system.

  select case(hdn_Option)
    
    ' case for the campus name
    
      Case "Campus" : ' code to edit the new campus into the database
      
  	                  ' Reteieving the contents from the input forms 
                         old_Campus = Request.form("hdnCampusID")
                         edit_Campus = Request.form("txtCampusName")
                       ' editing existing campus
                       
                         strSQL = "Select * from tblCampus where strCampusName = '"& edit_Campus &"'"
  	                      set rsCheckCampus = Server.CreateObject("ADODB.Recordset")
                          rsCheckCampus.Open strSQL, conn, 3, 3
                      
                          if rsCheckCampus.EOF = True then
                          
  					      strSQL ="Update tblCampus Set strCampusName ='"&edit_Campus&"' where numCampusID ="& old_Campus 
 					      set rsEditCampus = Server.CreateObject("ADODB.Recordset")
                          rsEditCampus.Open strSQL, conn, 3, 3 %> 
                           
                         <p><font color="#660033"><b>The Campus name has been edited successfully !</b></p> <%													
						 
						 'Redirect to admin functions page on success
						 response.redirect("admin.asp")                        
                         else
                         %>
                       <p><font color="#660033"><b>Record already exists ! 
								,&nbsp; Please click&nbsp; on the 'Back' button of the browser to enter the new 
								data.</b></p> 
                       <% end if
              
   ' case for the Building name
     Case "Building" : ' code to edit the new campus into the database
      
                        ' Reteieving the contents from the input forms 
                        
                           old_Building = Request.form("hdnBuildingId")
 						   edit_Building_Campus = Request.form("hdnCampusId")
						   edit_Building = Request.form("txtBuildingName")
						   						                                                      
                          strSQL = "Select * from tblBuilding where strBuildingName = '"& edit_Building &"' and numCampusID ="& edit_Building_Campus
  	                      set rsCheckBuilding = Server.CreateObject("ADODB.Recordset")
                          rsCheckBuilding.Open strSQL, conn, 3, 3
                      
                          if rsCheckBuilding.EOF = True then
      
                          strSQL ="Update tblBuilding Set strBuildingName ='"&edit_Building&"' where numCampusID ="& Edit_Building_Campus &" and numBuildingId = "& old_Building &""
 		 	              set rsEditCampus = Server.CreateObject("ADODB.Recordset")
                          rsEditCampus.Open strSQL, conn, 3, 3 %> 
                         
                         <p><font color="#660033"><b>The Building name has been edited successfully !</b></p> <%
                       
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
					   else %>
                       <p><font color="#660033"><b>Record already exists ! 
								,&nbsp; Please click&nbsp; on the 'Back' button of the browser to enter the new 
								data.</b></p> 
                       <%end if    
                       
  ' case for the Facility
     Case "Facility" : ' code to edit the new Facility into the database
      
                        ' Reteieving the contents from the input forms 
                          edit_Facility_Old_roomNumber = Request.Form("hdnRoomNumber")

                          edit_Facility_RoomNumber = Request.Form("txtRoomNumber")
						  edit_Facility_RoomName = Request.Form("txtRoomName")
						  'edit_Facility_Supervisor = Request.Form("cboSupervisorName")
						  edit_Facility_Building = Request.Form("hdnBuildingID")
						  edit_Facility_SupervisorID = Request.Form("cboSupervisorID")
                       
                       		'jan 2010 change to repair relationship facility:supervisor
                         '&"strFacilitySupervisor = '"& edit_Facility_Supervisor&"'"_  
                          strSQL ="Update tblFacility Set "_
                          &"strRoomNumber = '"&edit_Facility_RoomNumber&"',"_ 
                          &"strRoomName = '"&edit_Facility_RoomName&"',"_                                                  
                          &"numFacilitySupervisorID = '"& edit_Facility_SupervisorID&"'"_                                                 
                          &" where numBuildingId = "& edit_Facility_Building &" and strRoomNumber = '"& edit_Facility_Old_roomNumber&"'"
 					      set rsEditFacility = Server.CreateObject("ADODB.Recordset")
                          
                          'response.write(strSQL)
                          rsEditFacility.Open strSQL, conn, 3, 3 
                          
                          'get the numFacilityID form our recently updated record so that we can update the QORA table to reflect changes
                          strSQL = "select numFacilityID from tblFacility where numBuildingId = "&edit_Facility_Building&" "_
                          &" and strRoomNumber = '"&edit_Facility_RoomNumber&"'"
                          
                          set rsFacilityID = Server.CreateObject("ADODB.Recordset")
                          'Response.write(strSQL)
                          rsFacilityID.Open strSQL, conn, 3, 3 
                          facilityID = rsFacilityID("numFacilityID")
                          
                          //Get the string supervisor ID and update the tblQORA table
                          strSQL = "select strLoginID from tblFacilitySupervisor where numSupervisorID = "&edit_Facility_SupervisorID
                          set rsSupervisorStr = Server.CreateObject("ADODB.Recordset")
                          rsSupervisorStr.Open strSQL, conn, 3, 3 
                          edit_Facility_Supervisor = rsSupervisorStr("strLoginID")
                          
                                                    
                          strSQL = "Update tblQORA set strSupervisor = '"&edit_Facility_Supervisor &"' where numFacilityID = "&facilityID
                          set updateQORA = Server.CreateObject("ADODB.Recordset")
                          'Response.write(strSQL)
                          updateQORA.Open strSQL, conn, 3, 3 
                          %>
                          
                          <p><font color="#660033"><b>The Facility details have been edited successfully !</b></p><%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
	  Case "Operation":
	  					Dim operationID
	  					Dim operationName
	  					Dim supervisorID
	  					operationID = Request.form("hdnOperationId")
	  					operationName = Request.form("txtOperationName")
	  					supervisorID = Request.form("cboSupervisorID") 
	  					
	  					strSQL = "Update tblOperations set strOperationName = '"&operationName &"', numFacilitySupervisorID = "&supervisorID&" where numOperationID = "&operationID
                          set updateOperations = Server.CreateObject("ADODB.Recordset")
                          'Response.write(strSQL)
                          updateOperations.Open strSQL, conn, 3, 3 
                          
                          
                          //Get the string supervisor ID and update the tblQORA table
                          strSQL = "select strLoginID from tblFacilitySupervisor where numSupervisorID = "&supervisorID
                          set rsSupervisorStr = Server.CreateObject("ADODB.Recordset")
                          rsSupervisorStr.Open strSQL, conn, 3, 3 
                          dim superVisor
                          superVisor = rsSupervisorStr("strLoginID")
                          
                                                    
                          strSQL = "Update tblQORA set strSupervisor = '"&superVisor &"' where numOperationID = "&operationID
                          set updateQORA1 = Server.CreateObject("ADODB.Recordset")
                          'Response.write(strSQL)
                          updateQORA1.Open strSQL, conn, 3, 3 
                          %>
                          
                          <p><font color="#660033"><b>The Operation details have been edited successfully !</b></p><%
	  					
                          ' case for the Faculty name
      Case "Faculty" : ' code to edit the new Faculty into the database
      
  	                  ' Reteieving the contents from the input forms 
                         old_Faculty = Request.form("hdnFacultyId")
                         edit_Faculty = Request.form("txtFacultyName")
                         
                         edit_DGivenName = Request.form("txtDGivenName")
						 edit_DSurname =Request.form("txtDSurName")
						 edit_DLogin =Request.form("txtDLogin")
						 edit_DPassword=Request.form("txtDPassword")
						
						 
                       ' editing existing campus
                       
                         strSQL = "Select * from tblFaculty where strFacultyName = '"& edit_Campus &"'"
  	                      set rsCheckFaculty1 = Server.CreateObject("ADODB.Recordset")
                          rsCheckFaculty1.Open strSQL, conn, 3, 3
                      
                          if rsCheckFaculty1.EOF = true then
                          
  					      strSQL ="Update tblFaculty Set strFacultyName ='"&edit_Faculty&"',"_
  					              &"strDLogin = '"& edit_DLogin &"',"_
  					              &"strDPassword = '"& edit_DPassword &"',"_
  					              &"strDGivenName ='"& edit_DGivenName &"',"_
  					              &"strDSurname = '"& edit_DSurname &"'"_  
  					              &" where numFacultyID ="& old_Faculty 
  					              
 					      set rsEditFaculty = Server.CreateObject("ADODB.Recordset")
                          'response.write(strSQL)
                          rsEditFaculty.Open strSQL, conn, 3, 3 %> 
                           
                         <p><font color="#660033"><b>The Faculty / Unit name has been edited successfully !</b></p> <%
						 
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
                        
                         else
                         %>
                       <p><font color="#660033"><b>Record already exists ! 
								,&nbsp; Please click&nbsp; on the 'Back' button of the browser to enter the new 
								data.</b></p> 
                       <% end if

                      ' case for the Facility Supervisor
                      
      Case "Supervisor" : ' code to edit the Supervisor into the database

  	                  ' Reteieving the contents from the input forms 
  	                  
  	                  	 edit_Sup_ID = Request.form("hdnSupervisorId")
                         edit_Sup_newlogin = Request.form("txtnewID")
                         edit_Sup_SName = Request.form("txtSurName")
						 edit_Sup_GName = Request.form("txtGivenName")
						 edit_Sup_LoginId = Request.Form("hdnLoginId")
						 edit_Sup_Password = Request.Form("txtPassword")
						 edit_Sup_Faculty  = Request.Form("cboFaculty")
						 edit_Sup_Deprecated = Request.Form("deprecated")
                           'Response.write(edit_Sup_ID)                    
                       ' editing existing campus
                       
                          strSQL = "Select * from tblFacilitySupervisor where strLoginID = '"& edit_LoginID &"'"
  	                      set rsCheckFacilitySup = Server.CreateObject("ADODB.Recordset")
                          rsCheckFacilitySup.Open strSQL, conn, 3, 3
                          
                       'if the box is checked, then the record is active, i.e. not deprecated.
                       'Response.write(edit_Sup_Deprecated)
						  if edit_Sup_Deprecated = "on" then
   							 edit_Sup_Deprecated = "No"
						  else
  							 edit_Sup_Deprecated = "Yes"
						  end if
						  
						  'Here we wish to check if the newlogin ID already exists
						  strSQL ="Select * from tblFacilitySupervisor where strLoginID = '" & edit_Sup_newlogin &"'"
						  set rsCheckAlreadyExists = Server.CreateObject("ADODB.Recordset")
                          rsCheckAlreadyExists.Open strSQL, conn, 3, 3
                          
                          'if its the same as out existing user, then we can skip here (test for equality)
						  if not rsCheckAlreadyExists.eof and not (edit_Sup_LoginId = edit_Sup_newlogin) then
						  	
						  %> <script type="text/javascript">
							alert("Sorry, this login ID already exists, please choose another");
							location.href="admin.asp";
							
							</script><%
							'response.redirect("admin.asp")
						
                      	else
                          'otherwise all is well and we can proceed to update
  					      strSQL ="Update tblFacilitySupervisor Set strPassword ='"&edit_Sup_Password&"',"_
  					      &" strGivenName = '"& edit_Sup_GName &"',"_
  					      &" strSurname = '"& edit_Sup_SName&"',"_
						  &" numFacultyID = "& edit_Sup_Faculty&","_  
						  &" strLoginID = '"& edit_Sup_newlogin&"',"_ 
						  &" boolDeprecated =  "&edit_Sup_Deprecated	&""_		      
  					      &" where numSupervisorID = "& edit_Sup_ID &""
  					      
  					      'Response.write(strSQL)
 					      set rsEditFacility = Server.CreateObject("ADODB.Recordset")
                          rsEditFacility.Open strSQL, conn, 3, 3 
                          'response.write(strSQL) 
                          
                          strSQL = "Update tblQORA set strSupervisor ='"&edit_Sup_newlogin&"' where strSupervisor ='"&edit_Sup_LoginId&"'"
                          set updateQORA = Server.CreateObject("ADODB.Recordset")
                          updateQORA.Open strSQL, conn, 3, 3 
                          %>
                         <p><font color="#660033"><b>The supervisor details have been edited successfully !</b></p>
 <% 
 						end if
					   	'Redirect to admin functions page on success
						'response.redirect("admin.asp") 
						'- assure the user the trxn was completed %>
						
						<script type="text/javascript">
							alert("The supervisor details have been edited successfully !");
							location.href="admin.asp";
							
							</script>
							<%

                          
  end select    
  
  
%>

</body>
</html>