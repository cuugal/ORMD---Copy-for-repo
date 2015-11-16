
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
          // function to reload the form to add the new entries
          function FillDetailsSuper() {
              document.MenuSuper.submit();
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
                  data: "mode=" + "MenuOperation&numFacultyId="+numFacultyId+"&strSuperv="+strSuperv,
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
                          window.location.href = "/index.asp";
                      };
                      
                  }
              });
          }
          function FillDetailsTask() {
              document.MenuTask.submit();
          }

          function clearform() {
              var str
              str = "SearchQORA.asp";
              //window.location.replace(str);
              location.reload();
          }

          function ChangeType(val) {
              document.Form2.QORAtype.value = val;
              //console.log(document.Form2.QORAtype.value);

          }

          function fillSearch() {
          }2

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
                                 <select size="1" name="cboFacultySuper" tabindex="1" onchange="javascript:FillDetailsSuper()">
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
                           <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty" value="<%=cboFacultySuper%>" />
                           <input type="hidden" name="searchType" value="supervisor" />
                           <tr>
                              <th>Supervisor Name</th>
                              <td>
                                 <%'******* code to fill the Supervisor*****%>
                                 <%
                                    Dim connSup
                                    Dim rsFillSup
                                    Dim strSuperv
                                    
                                    'Database Connectivity Code 
                                      set connSup = Server.CreateObject("ADODB.Connection")
                                      connSup.open constr
                                     
                                     ' setting up the recordset
                                     
                                       strSQL ="Select * from tblFacilitySupervisor where numFacultyId ="&numFacultyId &" order by strGivenName "
                                       set rsFillSup = Server.CreateObject("ADODB.Recordset")
                                       rsFillSup.Open strSQL, connSup, 3, 3
                                                                        %>
                                 <%
                                    strSuperv = request.form("cboSupervisorName")
                                    
                                                                 %>
                                 <select size="1" name="cboSupervisorName" tabindex="2">
                                    <option value="0"
                                       <% if strSuperV = "" then
                                          response.Write "select any one"
                                          end if %>>Select any one</option>
                                    <%while not rsFillSup.Eof
                                       if rsFillSup("boolDeprecated") = 0 then%>
                                    <option value="<%=rsFillSup("strLoginID")%>"
                                       <% if rsFillSup("strLoginId") = strSuperV   then
                                          response.Write "selected"
                                          end if %>><%=cstr(rsFillSup("strGivenName")) + " " + cstr(rsFillSup("strsurname")) %></option>
                                    <% 
                                       End if
                                       rsFillSup.Movenext
                                       wend 
                                       
                                       ' closing the connections
                                       rsFillSup.close
                                       set rsFillSup = nothing
                                       connSup.Close
                                       set connSup = nothing
                                                               %>
                                 </select>
                                 &nbsp;
                              </td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /-->
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
                                 <select size="1" name="cboFacultyLocation" tabindex="1" onchange="javascript:FillDetailsLocation()">
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
                              <%'******* code to fill the Building*****%>
                              <%
                                 Dim conn
                                 Dim rsFillBuilding
                                 
                                 'Database Connectivity Code 
                                   set conn = Server.CreateObject("ADODB.Connection")
                                   conn.open constr
                                  
                                  ' setting up the recordset
                                  
                                        strSuperv = request.form("cboSupervisorName")
                                         numCampusID = cint(request.form("cboCampus"))
                                         'response.write(numCampusId)
                                        
                                  
                                    strSQL = "Select distinct(tblFacility.numBuildingId)as NumBuildingID,tblCampus.strCampusName,tblBuilding.strBuildingName "_
                                    &"from tblBuilding,tblCampus,tblFacility, tblFacilitySupervisor, tblFaculty "_
                                    &"where tblFaculty.numFacultyID="& numFacultyID&" "_
                                    &"and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultYID "_
                                    &"and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID "_
                                    &"and tblFacility.numBuildingId = tblBuilding.numBuildingId "_
                                    &"and tblBuilding.numCampusId = tblCampus.numCampusId "_
                                    &" order by strBuildingName"
                                    
                                    'response.write(strSQL)
                                  'end if
                                    
                                    set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
                                    rsFillBuilding.Open strSQL, conn, 3, 3
                                                     %>
                              <%    numBuildingID = cint(request.form("cboBuilding"))
                                 if numBuildingID = "" then
                                    numBuildingID = 0
                                 end if 
                                 
                                             %>
                              <td>
                                 <select size="1" name="cboBuilding" tabindex="4" onchange="javascript:FillDetailsLocation()">
                                    <option value="0"
                                       <% if numBuildingID = 0 then
                                          response.Write "select any one"
                                          end if %>>Select any one</option>
                                    <%while not rsFillBuilding.Eof%>
                                    <option value="<%=rsFillBuilding("numBuildingID")%>"
                                       <% if rsFillBuilding("numBuildingID") = numBuildingID then
                                          response.Write "selected"
                                          end if %>><%=cstr(rsFillBuilding("strBuildingName")) + " - " + cstr(rsFillBuilding("strCampusName")) + "  " + "Campus"%></option>
                                    <%rsFillBuilding.Movenext
                                       wend 
                                       
                                       ' closing the connections
                                       
                                         rsFillBuilding.close
                                         set rsFillBuilding = nothing
                                         conn.Close
                                         set conn = nothing
                                                          %>
                                 </select>
                              </td>
                           </tr>
                        </form>
                        <form method="post" name="Submit2" action="CollectInfo.asp" name="f1" enctype="application/x-www-form-urlencoded">
                           <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
                           <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
                           <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
                           <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
                           <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
                           <input type="hidden" name="cboFaculty" value="<%=cboFacultyLocation%>" />
                           <input type="hidden" name="searchType" value="location" />
                           <tr>
                              <th>Room No. / Name</th>
                              <%'******Code to fill the Room Name and Room Number****%>
                              <%
                                 Dim connR
                                 Dim rsFillR
                                 
                                 'Database Connectivity Code 
                                   set connR = Server.CreateObject("ADODB.Connection")
                                   connR.open constr
                                  
                                  ' setting up the recordset
                                  numCampusID = cint(request.form("cboCampus"))
                                  numBuildingID = cint(request.form("cboBuilding"))
                                 
                                    strSQL ="SELECT tblFacility.strRoomNumber,tblFacility.strRoomName,"_
                                    &" tblBuilding.strBuildingName,tblFacility.numFacilityId, strGivenName, strSurname"_
                                    &" FROM tblFacility, tblBuilding, tblFacilitySupervisor , tblFaculty"_ 
                                    &" WHERE tblFacility.numBuildingID=tblBuilding.numBuildingID "_
                                    &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
                                    &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_
                                    &" and tblFaculty.numFacultyID = "& numFacultyID&" "_
                                    &" And  tblBuilding.numBuildingId = "& numBuildingId &" "_
                                    &" order by tblFacility.strRoomName"
                                 
                                 
                                    set rsFillR = Server.CreateObject("ADODB.Recordset")
                                    rsFillR.Open strSQL, connR, 3, 3
                                                     %>
                              <td>
                                 <select size="1" name="cboRoom" tabindex="5">
                                    <option value="0">Select any one</option>
                                    <%While not rsFillR.EOF 
                                       if len(strSuperv) <= 1 then
                                          facility_name =cstr(rsFillR("strRoomNumber"))+ " - "+cstr(rsFillR("strRoomName"))&" - "&rsFillR("strGivenName")&" "&rsFillR("strSurname")
                                       else
                                          facility_name =cstr(rsFillR("strRoomNumber"))+ " - "+cstr(rsFillR("strRoomName"))
                                       end if	
                                                      %>
                                    <option value="<%=rsFillR("numFacilityId")%>"><%=facility_name%></option>
                                    <%
                                       rsFillR.Movenext
                                       wend
                                                    %>
                                 </select>
                              </td>
                           </tr>
                           <tr>
                              <td colspan="2">
                                 <center>
                                    <input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="    clearform()" />
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