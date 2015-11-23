
   <%
      if session("strLoginId") <> "admin" then
       response.redirect "AccessRestricted.htm"
      end if
      %>
   <%
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
          // function to reload the form to add the new entries
          function FillDetails() {
              document.SearchQORA.submit();
          }
          //Function to clear the contents of the form
          function resetForm() {
              document.Menu.txtHazardousTask.Value = "*"
          }
         
          function FillDetailsLocation() {
              document.MenuLocation.submit();
          }
          function FillDetailsOperation(numFacultyId, strSuperv) {
              $("#opsFacultyId").val(numFacultyId);
              // Fire off the request to /form.php
              request = $.ajax({
                  url: "AJAXSearch.asp",
                  type: "post",
                  data: "mode=" + "Operation&numFacultyId="+numFacultyId+"&strSuperv="+strSuperv,
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
                              $el.append($("<option></option>")
                                 .attr("value", value).text(key));
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
                  data: "mode=" + "Supervisor&numFacultyId=" + numFacultyId + "&strSuperv=" + strSuperv,
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
                              $el.append($("<option></option>")
                                 .attr("value", value).text(key));
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
                  data: "mode=" + "LocationBuilding&numFacultyId=" + numFacultyId + "&strSuperv=" + strSuperv,
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
                              $el.append($("<option></option>")
                                 .attr("value", value).text(key +" Campus"));
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
                              $el.append($("<option></option>")
                                 .attr("value", value).text(key + " Campus"));
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

   <body>
      <div id="wrapper" class="container">
         <div id="content">
            <h1 class="pagetitle">Search UTS Risk Assessments</h1>
            <center>
               <ul class="nav nav-tabs" style="width: 65%">
                  <li class="active"><a data-toggle="tab" href="#facility">Search Facility Locations</a></li>
                  <li><a data-toggle="tab" href="#operations">Search Operations/Projects</a></li>
                  <li><a data-toggle="tab" href="#supervisors">Search Supervisors</a></li>
                  <li><a data-toggle="tab" href="#ra">Search RA Number</a></li>
               </ul>
               <div class="tab-content" style="width: 65%">
                  <%'********************************** SEARCH SUPERVISOR  **************************************************************%>
                  <div id="supervisors" class="tab-pane fade">
                     <table class="adminfn">
                        <form method="post" action="SearchQORA.asp" name="MenuSuper">
                           <tr>
                              <td>Search Supervisor</td>
                           </tr>
                           <tr>
                              <th>Faculty/Unit</th>
                              <td>
                                 <%numFacultyID = cint(request.form("cboFacultySuper"))
                                    if numFacultyID = "" then
                                       numFacultyID = 0
                                    end if %>
                            
                                 <select size="1" autocomplete="off"  name="cboFacultySuper" tabindex="1" onchange="javascript:FillDetailsSupervisor(this.value, '<%=strsuperV%>')">
                                    <option value="0"
                                       <% if numFacultyID = 0 then
                                          response.Write "Select any one"
                                          end if %>>Select any one</option>
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
                        </form>
                        <form method="post" name="Submit1" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
                           <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
                           <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
                           <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
                           <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
                           <input type="hidden" name="hdnFacultyId" id="superFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty" value="<%=cboFacultySuper%>" />
                           <input type="hidden" name="searchType" value="supervisor" />
                           <tr>
                              <th>Supervisor Name</th>
                              <td>
                                                               
                                 <select size="1" name="cboSupervisorName" id="cboSupervisor" tabindex="2">
                                      <option value="0">Select any one</option>
                                    </select>
                                 &nbsp;
                              </td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
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
                        <form method="post" action="Menu.asp" name="MenuLocation">
                           <tr>
                              <td>Search Location</td>
                           </tr>
                           <tr>
                              <th>Faculty/Unit</th>
                              <td>
                                 <%    numFacultyID = cint(request.form("cboFacultyLocation"))
                                    if numFacultyID = "" then
                                       numFacultyID = 0
                                    end if %>
                                 <select size="1" autocomplete="off" name="cboFacultyLocation" id="cboFacultyLocation" tabindex="1" onchange="javascript:FillBuildingsLocation(this.value, '<%=strsuperV%>')">
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
                                 <select size="1" name="cboBuilding" id="cboBuilding" tabindex="4" onchange="javascript:FillRoomLocation(this.value, '<%=strsuperV%>')">
                                    <option value="0">Select any one</option>
         
                                 </select>
                              </td>
                           </tr>
                        </form>
                        <form method="post" name="Submit2" action="CollectInfo.asp" name="f1" enctype="application/x-www-form-urlencoded">
                           <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
                           <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
                           <input type="hidden" name="hdnBuildingId" id="locBuidingId" value="<%=numBuildingId%>" />
                           <input type="hidden" name="hdnCampusID" id="locCampusId" value="<%=numCampusId%>" />
                           <input type="hidden" name="hdnFacultyId" id="locFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty"  value="<%=cboFacultyLocation%>" />
                           <input type="hidden" name="searchType" value="location" />
                           <tr>
                              <th>Room No. / Name</th>
                             
                              <td>
                                 <select size="1" name="cboRoom" id="cboRoom" tabindex="5">
                                    <option value="0">Select any one</option>
                                    
                                 </select>
                              </td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />
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
                        <form method="post" action="SearchQORA.asp" name="MenuOperation">
                           <tr>
                              <td>Search Operation</td>
                           </tr>
                           <tr>
                              <th>Faculty/Unit</th>
                              <td>
                                 <%    numFacultyID = cint(request.form("cboFacultyOperation"))
                                    if numFacultyID = "" then
                                       numFacultyID = 0
                                    end if %>
                                 <select size="1" autocomplete="off"  name="cboFacultyOperation" tabindex="1" onchange="javascript:FillDetailsOperation(this.value, '<%=strsuperV%>')">
                                    <option value="0" " >Select any one</option>
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
                        
                        </form>
                        <form method="post" name="Submit3" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
                           <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
                           <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
                           <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
                           <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
                           <input type="hidden" name="hdnFacultyId" id="opsFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty" value="<%=cboFacultyOperation%>" />
                           <input type="hidden" name="searchType" value="operation" />
                           <tr>
                              <th>Operation</th>
                              <td>
                                 <select name="cboOperation" id="cboOperation">
                                    <option value="0">Select any one</option>
                                   
                                 </select>
                              </td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
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
                      <form method="post" action="SearchQORA.asp" name="MenuTask">
                        <table class="adminfn">
                       
                           <tr>
                              <td>Search Task/RA Number</td>
                           </tr>
                           <tr>
                              <th>Faculty/Unit</th>
                              <td>
                                 <%    numFacultyID = cint(request.form("cboFacultyTask"))
                                    if numFacultyID = "" then
                                       numFacultyID = 0
                                    end if %>
                                 <select size="1" name="cboFacultyTask" tabindex="1" onchange="javascript:FillDetailsTask()">
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
                        </form>
                        <form method="post" name="Submit4" action="CollectInfoAdmin.asp" name="f1" enctype="application/x-www-form-urlencoded">
                           <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
                           <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
                           <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
                           <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
                           <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty" value="<%=cboFacultyTask%>" />
                           <input type="hidden" name="searchType" value="task" />
                           <tr>
                              <th>Task/RA Number</th>
                              <td>
                                 <input type="text" name="txtHazardousTask" size="40" tabindex="0" />
                              </td>
                           </tr>
                           <tr>
                              <td></td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
                                    <!--DLJ Removed this button from common search 22July2011-->
                        
                        </center></td>
                        </form>
                        </tr>
                        <tr>
                           <td>&nbsp;</td>
                        </tr>
                     </table>
                  </div>
                  <%'************************************************  END TASK OPERATION ***************************************************** %>
               </div>
            </center>
         </div>
         <!-- close the content DIV -->
      </div>
      <!-- close the wrapper div -->