<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <style type="text/css">
 .navcontainer { width: 200px; font-size: 9pt; }

.navcontainer ul { margin-left: 0; padding-left: 0; list-style-type: none; font-family: Arial, Helvetica, sans-serif; }

.navcontainer a { display: block; padding: 3px;  width: 160px;  background-color: #EDF4F6; border: 1px solid #D7E9ED; margin-bottom: 2px;}

.navcontainer a:link, .navlist a:visited { color: #0083B3; text-decoration: none; margin-bottom: 2px;}

.navcontainer a:hover { background-color: #F8E9AD; color: #0083B3; margin-bottom: 2px;}
 </style>
 <title>Online Risk Register - Administration Functions</title>
</head>
<body>
<%
' Declaring the general variables
  Dim add_Campus
  Dim add_Building
  Dim add_Faculty
  Dim add_Supervisor
  
  dim add_DLogin
  dim add_DPassword
  dim add_DGivenName
  dim add_DSurname
  
  Dim add_RoomName
  Dim add_RoomNumber
  Dim add_Facility_Supervisor
  Dim add_Facility_Building
  
  Dim add_Operation
  
  Dim add_Building_Campus
  Dim hdn_Option
  
' Declaring the Database variables  
  
  Dim conn
  Dim RsCheckCampus
  Dim RsCheckBuilding
  Dim RsCheckFaculty
  Dim RsCheckSupervisor
  Dim RsCheckFacility
  Dim RsAddCampus
  Dim RsAddBuilding
  Dim RsAddFaculty
  Dim RsAddSupervisor
  Dim RsAddFacility
  Dim strSQL
  
  ' code for the database connectivity.
  
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
   
' Retrieving the contents from the input forms
 
  add_Campus = Request.form("txtCampusName")
  
  add_Building = Request.form("txtBuildingName")
  add_Building_Campus = Request.form("CboCampusName")
  
  add_Faculty = Request.form("txtFacultyName")
  add_DLogin = Request.form("txtDLoginID")
  add_DPassword = Request.form("txtDPassword")
  add_DGivenName = Request.form("txtDGivenName")
  add_DSurName = Request.form("txtDsurName")
  add_DPassword = Request.form("txtDPassword")
  
  
  add_Sup_Faculty = Request.form("cboFaculty")
  add_Sup_GName = Request.form("txtGivenName")
  add_Sup_SName = Request.form("txtSurname")
  add_Sup_LoginId = Request.form("txtLoginId")
  add_Sup_Password = Request.form("txtPassword")
  add_Sup_Type = Request.form("strAccessLevel")
  add_Sup_Email = Request.form("txtEmail")

  add_Operation = Request.form("txtOperationName")
  
  add_RoomName = Request.form("txtRoomName")
  add_RoomNumber = Request.form("txtRoomNumber")
  add_Facility_Supervisor = Request.form("cboSupervisorName")
  add_Facility_Building = Request.form("cboBuildingName")
  
  
  
  'strSQL = "Select * from tblFacilitySupervisor where strLoginId= '"& add_Facility_Supervisor &"'"
  
  'set rsAddFacultyId = Server.CreateObject("ADODB.Recordset")
  '    rsAddFacultyId.Open strSQL, conn, 3, 3
  
  'dim FacultyId
  'FacultyId = rsAddFacultyId("numFacultyId")
  
  hdn_Option = Request.Form("hdnOption")
             

   
' code to write the contents from the admin files into the database system.


  select case(hdn_Option)
    
    ' case for the campus name
      Case "Campus" : ' code to add the new campus into the database
  	                ' Checking the record into the database
  	                
  	                  strSQL = "Select * from tblCampus where strCampusName = '"& add_Campus &"'"
  	                  set rsCheckCampus = Server.CreateObject("ADODB.Recordset")
                      rsCheckCampus.Open strSQL, conn, 3, 3
                      
                      if rsCheckCampus.EOF = True then
                        ' adding a new record
 					        strSQL ="Insert Into tblCampus(strCampusName) Values ('"&add_Campus&"')"
 					        set rsAddCampus = Server.CreateObject("ADODB.Recordset")
                            rsAddCampus.Open strSQL, conn, 3, 3 %> 
                            
                             <p>The new Campus has been added successfully.</p> <%
							 
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")							 

                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p><%   
                      end if  
                      
  	
  	' case for the Building name
     
      Case "Building" : ' code to add the new Building into the database
  	                    ' Checking the record into the database
  	                
  	                  strSQL = "Select * from tblBuilding where strBuildingName = '"& add_Building &"' and numCampusId = "& add_Building_Campus &""
  	                  set rsCheckBuilding = Server.CreateObject("ADODB.Recordset")
                      rsCheckBuilding.Open strSQL, conn, 3, 3
                      
                      if rsCheckBuilding.EOF = True then
                        ' adding a new record
 					        strSQL ="Insert Into tblBuilding(strBuildingName,numCampusId) Values ('"&add_Building&"','"& add_Building_Campus&"')"
 					       set rsAddCampus = Server.CreateObject("ADODB.Recordset")
                          rsAddCampus.Open strSQL, conn, 3, 3
                            %> <p>The new Building has been added successfully.</p> <%
							
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
													
                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p> <%   
                      end if  
    
      	' case for the Facility/Unit 
     
      Case "Facility" : ' code to add the new Facility into the database
  	                    ' Checking the record into the database
  	                
  	                  strSQL = "Select * from tblFacility where numbuildingId = "& add_Facility_Building &" and strRoomName = '"& add_RoomName &"' and strRoomNumber = '"& add_RoomNumber &"'"
  	                  set rsCheckBuilding = Server.CreateObject("ADODB.Recordset")
                      rsCheckBuilding.Open strSQL, conn, 3, 3
                      
                      if rsCheckBuilding.EOF = True then
                        ' adding a new record
                        'AA jan 2010 relation fix altered
                        'strSQL ="Insert Into tblFacility(strRoomName,strRoomNumber,numBuildingId,strFacilitySupervisor) Values ('"&add_RoomName&"','"& add_Roomnumber &"','"& add_Facility_Building &"','"& add_Facility_Supervisor &"')"
 					     strSQL ="Insert Into tblFacility(strRoomName,strRoomNumber,numBuildingId,numFacilitySupervisorID) Values ('"&add_RoomName&"','"& add_Roomnumber &"','"& add_Facility_Building &"','"& add_Facility_Supervisor &"')"
 					        set rsAddCampus = Server.CreateObject("ADODB.Recordset")
                            rsAddCampus.Open strSQL, conn, 3, 3
                            %> <p>The new Facility has been added successfully.</p> <%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")							
							
                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p>
                      <%   
                      end if  
  
     ' case for the Faculty
     
     Case "Operation" : ' code to add the new Operation into the database
  	                    ' Checking the record into the database
  	                
  	                  strSQL = "Select * from tblOperations where strOperationName = '"& add_Operation &"'"
  	                  set rsCheckBuilding = Server.CreateObject("ADODB.Recordset")
                      rsCheckBuilding.Open strSQL, conn, 3, 3
                      
                      if rsCheckBuilding.EOF = True then
                        ' adding a new record
 					     strSQL ="Insert Into tblOperations(numFacilitySupervisorId, strOperationName) Values ("&add_Facility_Supervisor&",'"& add_Operation &"')"
 					        set rsAddCampus = Server.CreateObject("ADODB.Recordset")
                            rsAddCampus.Open strSQL, conn, 3, 3
                            %> <p>The new Facility has been added successfully.</p> <%

					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")							
							
                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p>
                      <%   
                      end if  
  
     ' case for the Faculty
     
       Case "Faculty" : ' code to add the new campus into the database
  	                ' Checking the record into the database
  	                
  	                  strSQL = "Select * from tblFaculty where strFacultyName = '"& add_Faculty &"' or strDLogin = '"& add_DLogin&"'"
  	                  set rsCheckFaculty= Server.CreateObject("ADODB.Recordset")
                      rsCheckFaculty.Open strSQL, conn, 3, 3
                      
                      if rsCheckFaculty.EOF = True then
                        ' adding a new record
 					        strSQL ="Insert Into tblFaculty(strFacultyName,strDgivenName,strDSurname,strDLogin,strDPassword) Values ('"&add_Faculty&"','"&add_DGivenName&"','"&add_DSurname&"','"&add_Dlogin&"','"&add_DPassword&"')"
 					        set rsAddFaculty = Server.CreateObject("ADODB.Recordset")
                            rsAddFaculty.Open strSQL, conn, 3, 3 %> 
                            
                             <p>The new Faculty/Unit has been added successfully.</p> <%
							 
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
													 
                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p> <%   
                      end if  
     
      ' case for the Supervisor
     
       Case "Supervisor" : ' code to add the new campus into the database
  	                ' Checking the LoginId record into the database
  	                
                      strSQL = "Select * from tblFacilitySupervisor where strLoginId = '"& add_Sup_LoginId &"'"
  	                  set rsCheckFacilitySup= Server.CreateObject("ADODB.Recordset")
                      rsCheckFacilitySup.Open strSQL, conn, 3, 3
                      
                      if rsCheckFacilitySup.EOF = True then
                        ' adding a new record
 					        strSQL ="Insert Into tblFacilitySupervisor(numFacultyID,strSurname,strGivenName,strLoginId,strPassword,strAccessLevel) Values ('"&add_sup_Faculty&"','"&add_sup_Sname&"','"&add_sup_Gname&"','"&add_sup_LoginId&"','"&add_sup_Password &"','"&add_Sup_Type&"')"
 					        set rsAddFacilitySupervisor = Server.CreateObject("ADODB.Recordset")
                            rsAddFacilitySupervisor.Open strSQL, conn, 3, 3 %> 
                            
                             <p>The new Supervisor has been added successfully.</p> <%
							 
					   	'Redirect to admin functions page on success
						response.redirect("admin.asp")
													 
                      else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p>
                      <%
                                            end if

           Case "Register" : ' code to add the new user into the database
          	                ' Checking the LoginId record into the database

                              strSQL = "Select * from tblFacilitySupervisor where strLoginId = '"& add_Sup_LoginId &"'"
          	                  set rsCheckFacilitySup= Server.CreateObject("ADODB.Recordset")
                              rsCheckFacilitySup.Open strSQL, conn, 3, 3

                              if rsCheckFacilitySup.EOF = True then
                                ' adding a new record
         					        strSQL ="Insert Into tblFacilitySupervisor(numFacultyID,strSurname,strGivenName,strLoginId,strPassword,strAccessLevel, strEmail) Values ('"&add_sup_Faculty&"','"&add_sup_Sname&"','"&add_sup_Gname&"','"&add_sup_LoginId&"','"&add_sup_Password &"','"&add_Sup_Type&"','"&add_Sup_Email&"')"
         					        set rsAddFacilitySupervisor = Server.CreateObject("ADODB.Recordset")
                                    rsAddFacilitySupervisor.Open strSQL, conn, 3, 3 %>

                                     <p>The new Supervisor has been added successfully.</p> <%

        					   	'Redirect to admin functions page on success
        						response.redirect("Home.asp")

                              else %><p>This record already exists. Please click the 'Back' button of the browser and re-enter.</p>

<div align="center">&nbsp;</div><%   
                      end if  


  end select    
  
%>

</body>
</html>