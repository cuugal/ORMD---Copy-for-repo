
<link rel="stylesheet" type="text/css" href="DataTables-1.10.10/css/dataTables.bootstrap.css" media="all" />
<script src="DataTables-1.10.10/js/jQuery.dataTables.min.js"></script>
<script src="DataTables-1.10.10/js/dataTables.foundation.min.js"></script>
<script src="DataTables-1.10.10/js/dataTables.bootstrap.min.js"></script>


   <%
       'Empty out session
        Session("hdnHTask") = ""
        Session("hdnBuildingId") = ""
        Session("hdnCampusId") = ""
        Session("hdnFacultyId") = ""
        Session("hdnFacilityId") = ""
        Session("hdnSupervisor") = ""
        Session("hdnOperationID") = ""
        Session("searchType") = ""

       session("cboOperation") = 0
	   session("cboFacility") = 0

	  Dim connFac
	  Dim rsFillFac
	  Dim strSQL
	  
	  'Database Connectivity Code 
		set connFac = Server.CreateObject("ADODB.Connection")
		connFac.open constr
	   
	   ' setting up the recordset
	   
		 strSQL ="Select * from tblFaculty order by strFacultyName"
		 set rsFillFac = Server.CreateObject("ADODB.Recordset")
		 rsFillFac.Open strSQL, connFac, 3, 3

		numCampusID = cint(request.form("cboCampus"))
	  %>
	  <script type="text/javascript">
		  // function to ask about the confirmation of the file.
		  function ConfirmChoice() {
			  answer = confirm("Are you sure you want to save this record?")
			  if (answer == true) {
				  return;
			  }
			  else {
				  return (false);
			  }
		  }

		  function FillDetailsOperation(numFacultyId, strSuperv) {
			  $("#opsFacultyId").val(numFacultyId);
			  // Fire off the request to /form.php
			  request = $.ajax({
				  url: "AJAXSearch.asp",
				  type: "post",
				  data: "mode=Operation&numFacultyId="+numFacultyId+"&strSuperv="+strSuperv,
				  async: false,
				  success: function (data) {
					  var jsonResult;
					  try {
						  var obj = jQuery.parseJSON(data);
						  var newOptions = obj.result;
						  var $el = $("#cboOperation");
						  $el.empty(); // remove old options
						  $el.append($("<option></option>").attr("value", 0).text("Select any one"));

						  $.each(newOptions, function (value, key) {
						      $.each(key, function(value, key){
						          $el.append($("<option></option>")
                                     .attr("value", value).text(key));
						      })
						  });
						  
					  }
					  catch (e) {
						  window.location.href = "/menu.asp";
					  };
					  
				  }
			  });
		  }
		  function FillDetailsSupervisor(numFacultyId, strSuperv) {
			  $("#superFacultyId").val(numFacultyId);
			  // Fire off the request to /form.php
			  request = $.ajax({
				  url: "AJAXSearch.asp",
				  type: "post",
				  data: "mode=Supervisor&numFacultyId=" + numFacultyId + "&strSuperv=" + strSuperv,
				  async: false,
				  success: function (data) {
					  var jsonResult;
					  try {
						  var obj = jQuery.parseJSON(data);
						  var newOptions = obj.result;
						  var $el = $("#cboSupervisor");
						  $el.empty(); // remove old options
						  $el.append($("<option></option>").attr("value", 0).text("Select any one"));
						  $.each(newOptions, function (value, key) {
						      $.each(key, function (value, key) {
						          $el.append($("<option></option>")
                                     .attr("value", value).text(key));
						      })
						  });
					  }
					  catch (e) {
						  window.location.href = "/menu.asp";
					  };

				  }
			  });
		  }

		  function FillBuildingsLocation(numFacultyId, strSuperv) {
			  $("#locFacultyId").val(numFacultyId);
			  // Fire off the request to /form.php
			  request = $.ajax({
				  url: "AJAXSearch.asp",
				  type: "post",
				  data: "mode=LocationBuilding&numFacultyId=" + numFacultyId + "&strSuperv=" + strSuperv,
				  async: false,
				  success: function (data) {
					  var jsonResult;
					  try {
						  var obj = jQuery.parseJSON(data);
						  var newOptions = obj.result;
						  var $el = $("#cboBuilding");
						  $el.empty(); // remove old options
						  $el.append($("<option></option>").attr("value", 0).text("Select any one"));
						  $.each(newOptions, function (value, key) {
						      $.each(key, function (value, key) {
						          $el.append($("<option></option>")
                                     .attr("value", value).text(key));
						      })
						  });
					  }
					  catch (e) {
						  window.location.href = "/menu.asp";
					  };

				  }
			  });
		  }
		  
		  function FillRoomLocation(numBuildingId, strSuperv) {
			  $("#locBuidingId").val(numBuildingId);
			  var numFacultyId = $("#cboFacultyLocation").val();
			  // Fire off the request to /form.php
			  request = $.ajax({
				  url: "AJAXSearch.asp",
				  type: "post",
				  data: "mode=" + "LocationRoom&numBuildingId="+numBuildingId +"&numFacultyId=" + numFacultyId + "&strSuperv=" + strSuperv,
				  async: false,
				  success: function (data) {
					  var jsonResult;
					  try {
						  var obj = jQuery.parseJSON(data);
						  var newOptions = obj.result;
						  var $el = $("#cboRoom");
						  $el.empty(); // remove old options
						  $el.append($("<option></option>").attr("value", 0).text("Select any one"));
						  $.each(newOptions, function (value, key) {
						      $.each(key, function (value, key) {
						          $el.append($("<option></option>")
                                     .attr("value", value).text(key));
						      })
						  });
					  }
					  catch (e) {
						  window.location.href = "/menu.asp";
					  };

				  }
			  });
		  }

		  function clearform() {
			  var str
	 
			  location.reload();
		  }


	  </script>

                
			   <ul class="nav nav-tabs" >
                 
				  <li class="active"><a data-toggle="tab" href="#facility">Facility Locations</a></li>
				  <li><a data-toggle="tab" href="#operations">Operations/Projects</a></li>
				  <li><a data-toggle="tab" href="#supervisors">Supervisors</a></li>
				  <li><a data-toggle="tab" href="#ra">Keyword Search</a></li>
             
				   <% if session("LoggedIn")= true then %>
				   <li><a data-toggle="tab" href="#my"><font color="#E60A0A">My Risk Assessments</font></a></li>
				   <% end if %>
                  
			   </ul>
			   <div class="tab-content">

               
				  <%'********************************** SEARCH SUPERVISOR  **************************************************************%>
				  <div id="supervisors" class="tab-pane fade">
					 <table class="adminfn">
						  
						   <tr>
							  <th>Faculty/Unit</th>
							  <td>
								 <%numFacultyID = cint(request.form("cboFacultySuper"))
									if numFacultyID = "" then
									   numFacultyID = 0
									end if %>
							
								 <select size="1" autocomplete="false" class="form-control" name="cboFacultySuper" id="cboFacultySuper" tabindex="1" onchange="javascript:FillDetailsSupervisor(this.value, '<%=strsuperV%>')">
									<option value="0">Select any one</option>
									<%while not rsFillFac.Eof 
									   if rsFillFac("boolActive")= True Then %>
									<option value="<%=rsFillFac("NumFacultyID")%>"
									   <% if rsFillFac("NumFacultyID") = numFacultyID Then
										  response.Write "selected"
										  end if %>><%=cstr(rsFillFac("strFacultyName"))%></option>
									<% End If
									   rsFillFac.Movenext	
									   wend 
									   
															   %>
								 </select>
							  </td>
						   </tr>
						
						<form method="post" autocomplete="false"  name="Submit1" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
						   <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
						   <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
						   <input type="hidden" name="hdnBuildingId" value="0" />
						   <input type="hidden" name="hdnCampusID" value="0" />
						   <input type="hidden" name="hdnFacultyId" id="superFacultyId" value="0" />
						   <input type="hidden" name="cboFaculty" value="<%=cboFacultySuper%>" />
						   <input type="hidden" name="searchType" value="supervisor" />
						   <tr>
							  <th>Supervisor Name</th>
							  <td>
															   
								 <select autocomplete="false" class="form-control" size="1" name="cboSupervisorName" id="cboSupervisor" tabindex="2">
									  <option value="0">Select any one</option>
									</select>
								 &nbsp;
							  </td>
						   </tr>
						   <tr>
							  <td colspan="2">
								 <center>
									<input type="Submit" class="btn btn-primary" value="Search" name="btnSearch" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" class="btn btn-primary" value="Clear Form" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
									<!--DLJ Removed this button from common search 22July2011-->
						</form>
						</center>
						</td>
						</tr>
						<tr>
						   <td>&nbsp;</td>
						</tr>
					 </table>
				  </div>
				  <%'************************************************ END SEARCH SUPERVISOR ******************************************************** %>
				  <%'************************************************  SEARCH LOCATION ******************************************************** %>
				  <div id="facility" class="tab-pane fade in active">
					 <table class="adminfn">
						 
						   <tr>
							  <th>Faculty/Unit</th>
							  <td>
								 <%    numFacultyID = cint(request.form("cboFacultyLocation"))
									if numFacultyID = "" then
									   numFacultyID = 0
									end if %>
								 <select size="1" autocomplete="false" class="form-control" name="cboFacultyLocation" id="cboFacultyLocation" tabindex="1" onchange="javascript:FillBuildingsLocation(this.value, '<%=strsuperV%>')">
									<option value="0"
									   <% if numFacultyID = 0 then
										  response.Write "Select any one"
										  end if %>>Select any one</option>
									<%rsFillFac.MoveFirst
									   while not rsFillFac.Eof 
											   'DLJ put this if statement in 22 Jan 2010 - is this OK?
											   if rsFillFac("boolActive")= True Then %>
									<option value="<%=rsFillFac("NumFacultyID")%>"
									   <% if rsFillFac("NumFacultyID") = numFacultyID Then
										  response.Write "selected"
										  end if %>><%=cstr(rsFillFac("strFacultyName"))%></option>
									<% End If
									   rsFillFac.Movenext	
									   wend 
									   
											   %>
								 </select>
							  </td>
						   </tr>
						   <tr>
							  <th>Building</th>
							 
								 
										   
							  <td>
								 <select autocomplete="false" class="form-control" size="1" name="cboBuilding" id="cboBuilding" tabindex="4" onchange="javascript:FillRoomLocation(this.value, '<%=strsuperV%>')">
									<option value="0">Select any one</option>
		 
								 </select>
							  </td>
						   </tr>
						
						<form method="post" autocomplete="false"  name="Submit2" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
						   <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
						   <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
						   <input type="hidden" name="hdnBuildingId" id="locBuidingId" value="0" />
						   <input type="hidden" name="hdnCampusID" id="locCampusId" value="0" />
						   <input type="hidden" name="hdnFacultyId" id="locFacultyId" value="0" />
						   <input type="hidden" name="cboFaculty"  value="0" />
						   <input type="hidden" name="searchType" value="location" />
						   <tr>
							  <th>Room No. / Name</th>
							 
							  <td>
								 <select autocomplete="false" class="form-control" size="1" name="cboRoom" id="cboRoom" tabindex="5">
									<option value="0">Select any one</option>
									
								 </select>
							  </td>
						   </tr>
						   <tr>
							  <td colspan="2">
								 <center>
									<input type="Submit" class="btn btn-primary" value="Search" name="btnSearch" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" class="btn btn-primary" value="Clear Form" name="btnClear" onclick="    clearform()" />
									&nbsp;&nbsp;&nbsp;&nbsp;
						</form>
						</center></td>
						</tr>
						
					 </table>
				  </div>
				  <%'************************************************ END SEARCH LOCATION ******************************************************** %>
				  <%'************************************************  SEARCH OPERATION ******************************************************** %>
				  <div id="operations" class="tab-pane fade">
					 <table class="adminfn">
						
						   <tr>
							  <th>Faculty/Unit</th>
							  <td>
								 <%    numFacultyID = cint(request.form("cboFacultyOperation"))
									if numFacultyID = "" then
									   numFacultyID = 0
									end if %>
								 <select size="1" class="form-control" autocomplete="false"  name="cboFacultyOperation" id="cboFacultyOperation" tabindex="1" onchange="javascript:FillDetailsOperation(this.value, '<%=strsuperV%>')">
									<option value="0" >Select any one</option>
									<%rsFillFac.MoveFirst
									   while not rsFillFac.Eof 
											   'DLJ put this if statement in 22 Jan 2010 - is this OK?
											   if rsFillFac("boolActive")= True Then %>
									<option value="<%=rsFillFac("NumFacultyID")%>"
									  ><%=cstr(rsFillFac("strFacultyName"))%></option>
									<% End If
									   rsFillFac.Movenext	
									   wend 
									   
								   %>
								 </select>
							  </td>
						   </tr>
						
						<form method="post" autocomplete="false"  name="Submit3" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
						   <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
						   <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
						   <input type="hidden" name="hdnBuildingId" value="0" />
						   <input type="hidden" name="hdnCampusID" value="0" />
						   <input type="hidden" name="hdnFacultyId" id="opsFacultyId" value="0" />
						   <input type="hidden" name="cboFaculty" value="0" />
						   <input type="hidden" name="searchType" value="operation" />
						   <tr>
							  <th>Operation</th>
							  <td>
								 <select autocomplete="false" class="form-control" name="cboOperation" id="cboOperation">
									<option value="0">Select any one</option>
								   
								 </select>
							  </td>
						   </tr>
						   <tr>
							  <td colspan="2">
								 <center>
									<input type="Submit" class="btn btn-primary" value="Search" name="btnSearch" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" class="btn btn-primary" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
									<!--DLJ Removed this button from common search 22July2011-->
						</form>
						</center></td>
						</tr>
						<tr>
						   <td>&nbsp;</td>
						</tr>
					 </table>
				  </div>
				  <%'************************************************ END SEARCH OPERATION ******************************************************** %>
				  <%'************************************************  SEARCH TASK ******************************************************** %>
				  <div id="ra" class="tab-pane fade">
					 
						<table class="adminfn">
					   
						  
						 
						
						<form method="post" autocomplete="false"  name="Submit4" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
						   <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
						   <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
						   <input type="hidden" name="hdnBuildingId" value="0" />
						   <input type="hidden" name="hdnCampusID" value="0" />
						   <input type="hidden" name="hdnFacultyId" value="0" />
						   <input type="hidden" name="cboFaculty" value="0" />
						   <input type="hidden" name="searchType" value="task" />
						   <tr>
							  <th>RA Number/Task</th>
							  <td>
								 <input type="text" class="form-control" id="txtHazardousTask" name="txtHazardousTask" size="40" tabindex="0" />
							  </td>
						   </tr>
						   <tr>
							  <td></td>
						   </tr>
						   <tr>
							  <td colspan="2">
								 <center>
									<input type="Submit" class="btn btn-primary" value="Search" id="taskSearch" name="btnSearch"  />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" class="btn btn-primary" value="Clear Form" name="btnClear" onclick="clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;
									<!--DLJ Removed this button from common search 22July2011-->
						
						</center></td>
						</form>
                            <script type="text/javascript">
                                $('#taskSearch').click(function () {
                                    if ($.trim($('#txtHazardousTask').val()) == '') {
                                        alert('RA Number/Task can not be left blank');
                                        return false;
                                    }
                                    return true;
                                });
                            </script>
						</tr>
						<tr>
						   <td>&nbsp;</td>
						</tr>
					 </table>
				  </div>
				  <%'************************************************  END TASK OPERATION ***************************************************** %>
                 
				  <%'************************************************  START MY RA OPERATION ***************************************************** %>
				    <% if session("LoggedIn") then 
						  Dim connFaci
						  Dim rsFillFaci
						  Dim strSQLFaci
	  
						  'Database Connectivity Code 
						  set connFaci = Server.CreateObject("ADODB.Connection")
						  connFaci.open constr
	   
						   ' setting up the recordset
	   
							strSQLFaci ="Select * from tblFacility "

							if session("numSupervisorId") <> 1 then
								strSQLFaci =strSQLFaci&" WHERE numFacilitySupervisorId = "& session("numSupervisorId")
							end if
							strSQLFaci = strSQLFaci&" order by strRoomNumber"
						   set rsFillFaci = Server.CreateObject("ADODB.Recordset")
						   rsFillFaci.Open strSQLFaci, connFaci, 3, 3

						  Dim connProj
						  Dim rsFillProj
						  Dim strSQLProj
	  
						  'Database Connectivity Code 
						  set connProj = Server.CreateObject("ADODB.Connection")
						  connProj.open constr
	   
						   ' setting up the recordset
	   
							strSQLProj ="Select * from tblOperations "

							if session("numSupervisorId") <> 1 then
								strSQLProj =strSQLProj&" WHERE numFacilitySupervisorId = "& session("numSupervisorId")
							end if
							strSQLProj = strSQLProj&" order by strOperationName"
						   set rsFillProj = Server.CreateObject("ADODB.Recordset")
						   rsFillProj.Open strSQLProj, connProj, 3, 3

						'response.write(strSQLProj)
						 %>
					  <div id="my" class="tab-pane fade">
                        
						<form method="post" autocomplete="false" action="SupRDateModified.asp"  name="FormA" enctype="application/x-www-form-urlencoded">
							  

                                        <script type="text/javascript">
                                            function checkFacility() {
                                                if ($("#myfacility").val() == 0) {
                                                    alert("Please select a Facility");
                                                    return false;
                                                }
                                                return true;
                                            }
                                            function checkOperation() {
                                                if ($("#myoperation").val() == 0) {
                                                    alert("Please select an Operation");
                                                    return false;
                                                }
                                                return true;
                                            }
                                        </script>

					       <table class="adminfn">
						   
							<tr>
								<th>Faculty/Unit:</th>
								<td colspan ="3"><strong><% =session("strFacultyName") %></strong></td>
							</tr>
							<tr>
								<th>Supervisor Name:</th>
								<td colspan="3"><strong><% =session("strName") %></strong></td>
							</tr>
					
 
							<input type="hidden" name="searchType" ID="searchType" value=""/>
						   <tr>
							  <th>Facility</th>
							  <td colspan="2">
								
								 <select autocomplete="false" class="form-control" id="myfacility" size="1" name="cboFacility" tabindex="1" onchange="$('#searchType').val('location')">
									<option value="0">Select any one</option>
									<%
									   while not rsFillFaci.Eof
											concat = rsFillFaci("strRoomNumber")&" "&rsFillFaci("strRoomName")
											 %>   
											<option value="<%=rsFillFaci("numFacilityId")%>"><%=concat%></option>
									<% 
										rsFillFaci.Movenext	
									   wend 
									%>
								 </select>
							  </td>
                               <td><input type="submit" size="70" value="List Facility Risk Assessments" class="btn btn-primary" onclick="return checkFacility();" name="btnGenRep" /></td>
						   </tr>
						
						<tr><td colspan="4"><hr /></td></tr>
						   <tr>
							  <th>Operation/Project</th>
							  <td colspan="2">
								 <select autocomplete="false" class="form-control" id="myoperation" name="cboOperation" id="cboOperation" Onchange="$('#searchType').val('operation')">
									<option value="0">Select any one</option>
								   <%
									   while not rsFillProj.Eof
											 %>   
											<option value="<%=rsFillProj("numOperationId")%>">
												<%=rsFillProj("strOperationName")%></option>
									<% 
										rsFillProj.Movenext	
									   wend 
									%>
								 </select>
							  </td>
						   
							  <td>
                                   <input type="submit" size="70"value="List Operation/Project Risk Assessments" class="btn btn-primary" onclick="return checkOperation();" name="btnGenRep" /></td>
						   </tr>
						   

                               <tr><td colspan="4"><hr /></td></tr>
                               <tr><td colspan="3"></td><td>
                               <button class="btn btn-primary" type="button" onclick="$('#myfacility').val(0);$('#myoperation').val(0);">Clear</button>
                                   </td></tr>
                               <tr><td colspan="3"></td><td>
                                   <button  class="btn btn-primary" type="button" onclick="window.location='LocationSup.asp'">Create New Risk Assessment</button>
                               
                              </td>
                               </tr>
						<tr>
						   <td>&nbsp;</td>
						</tr>
					 </table>
                          </form>
				  </div>
				   <% end if %>
				   <%'************************************************  END MY RA OPERATION ***************************************************** %>
			   </div>


<%
    set con	= server.createobject ("adodb.connection")
    con.open constr

    strSQL = "SELECT tblfaculty.strFacultyName, 'Facility' as searchType, tblFaculty.numFacultyID as key_id, count(numQORAId) as numRA, sum(iif(dtReview > Date() , 1 , 0 )) as numCurr  "_
    &"from tblFaculty, tblfacilitysupervisor, tblfacility, tblQORA "_
    &"where tblfaculty.numfacultyID = tblfacilitysupervisor.numfacultyid "_
    &"and tblfacility.numfacilitysupervisorid = tblfacilitysupervisor.numsupervisorid "_
    &"and tblqora.numfacilityid = tblfacility.numfacilityid "_
    &"group by strFacultyName, tblFaculty.numFacultyID "_

    &"union all "_

    &"SELECT tblfaculty.strFacultyName, 'Operation' as searchType, tblFaculty.numFacultyID as key_id, count(numQORAId) as numRA, sum(iif(dtReview > Date() , 1 , 0 )) as numCurr "_
    &"from tblFaculty, tblfacilitysupervisor, tbloperations, tblQORA "_
    &"where tblfaculty.numfacultyID = tblfacilitysupervisor.numfacultyid "_
    &"and tbloperations.numfacilitysupervisorid = tblfacilitysupervisor.numsupervisorid "_
    &"and tbloperations.strOperationName not like 'Archive%' "_
    &"and tblqora.numoperationid = tbloperations.numoperationid "_

    &"group by strFacultyName, tblFaculty.numFacultyID "_

    &"order by  2, strFacultyName"

    set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
    rsFillFaculty.Open strSQL, con, 3, 3



    set con	= server.createobject ("adodb.connection")
    con.open constr


    strSQL = "SELECT count(numQORAId) as numRA, 'location' as searchType, tblFacility.numFacilityId as key_id, sum(iif(dtReview > Date() , 1 , 0 )) as numCurr, strRoomName&' '&strRoomNumber as location "_
    &" FROM tblFacility,tblQORA "_
    &" Where tblQORA.numFacilityID = tblFacility.numFacilityID  group by strRoomName, strRoomNumber, tblFacility.numFacilityId"_
 
 
 &" union all "_
 
 &"SELECT count(numQORAId) as numRA, 'operation' as searchType, tblOperations.numOperationId as key_id, sum(iif(dtReview > Date() , 1 , 0 )) as numCurr, strOperationName as location "_
 &" FROM tblOperations ,tblQORA "_
 &" where tblQORA.numOperationID = tblOperations.numOperationId and strOperationName not like 'Archive%' group by strOperationName, tblOperations.numOperationId"
    
    ' response.write strSQL
    set rsFillOperation = Server.CreateObject("ADODB.Recordset")
        rsFillOperation.Open strSQL, con, 3, 3
   
     %>
<br />
<h3 style="float:left">UTS Risk Assessment Overview</h3>

<div style="clear:both"></div>
<ul class="nav nav-tabs" >
                 
	<li class="active"><a data-toggle="tab" href="#faculty_list">By Faculty</a></li>
	<li><a data-toggle="tab" href="#fac_oper_list">By Facility/Operation</a></li>

                  
</ul>
<div class="tab-content" style="padding:5px;">
    <div id="faculty_list" class="tab-pane fade in active">
     <table id="factable" class="display" cellspacing="0" width="100%">
        <thead>
            <tr>
                <th>Faculty Name</th>
                <th>Facility/Operation</th>
                <th>Current RA</th>
                <th>Total RA</th>
                
                <th>% Current</th>
            </tr>
        </thead>

        <tbody>
         <% dim link
             while not rsFillFaculty.Eof %>
            <tr>
                <% 
                    if rsFillFaculty("searchType") = "Facility" then
                        link = "CollectInfoAdmin.asp?searchType=location&hdnFacultyId="&rsFillFaculty("key_id") 
                    else
                        link = "CollectInfoAdmin.asp?searchType=operation&hdnFacultyId="&rsFillFaculty("key_id") 
                    end if
                    %>
                <td><a style="text-decoration: underline;" href="<%=link %>"><%=rsFillFaculty("strFacultyName") %></a></td>
                <td><%=rsFillFaculty("searchType") %></td>
                 <td><%=rsFillFaculty("numCurr") %></td>
                <td><%=rsFillFaculty("numRA") %></td>
               
                <td><%=formatnumber( cint(rsFillFaculty("numCurr"))/cint(rsFillFaculty("numRA")) *100,2)%></td>
            </tr>
            <%
            rsFillFaculty.Movenext
            wend
                 %>
            </tbody>
        </table>
    </div>
    <div id="fac_oper_list" class="tab-pane fade">
    <table id="ratable" class="display" cellspacing="0" width="100%">
        <thead>
            <tr>
                <th>Facility Name/Number or Operation</th>
                <th>Current RA</th>
                <th>Total RA</th>
                
                <th>% Current</th>
            </tr>
        </thead>

        <tbody>
         <% 
             while not rsFillOperation.Eof %>
            <tr>
                <% 
                    if rsFillOperation("searchType") = "location" then
                        link = "CollectInfoAdmin.asp?searchType=location&cboroom="&rsFillOperation("key_id") 
                    else
                        link = "CollectInfoAdmin.asp?searchType=operation&cboOperation="&rsFillOperation("key_id") 
                    end if
                    %>
                <td><a style="text-decoration: underline;" href="<%=link %>"><%=rsFillOperation("location") %></a></td>
                 <td><%=rsFillOperation("numCurr") %></td>
                <td><%=rsFillOperation("numRA") %></td>
               
                <td><%=formatnumber( cint(rsFillOperation("numCurr"))/cint(rsFillOperation("numRA")) *100,2)%></td>
            </tr>
            <%
            rsFillOperation.Movenext
            wend
                 %>
            </tbody>
        </table>

   </div>
 </div>
		  <script type="text/javascript">
		      $(document).ready(function () {
		          $('#ratable').DataTable();
		          $('#factable').DataTable();
		          $("form").each(function () {
		            
		              $(this).trigger("reset");
		          });
		          $("#cboFacultyLocation").val(0);
		          $("#cboFacultyOperation").val(0);
		          $("#cboFacultySuper").val(0);
		      });

    </script>