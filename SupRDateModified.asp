<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
dim loginId
loginId = session("strLoginId")
'Response.Write(loginId)
%>
<head>
    <!--#include file="bootstrap.inc"--> 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-au" />

<title>Online Risk Register - UTS Risk Assessments</title>

<script type="text/javascript" src="sorttable.js"></script>
</head>
<%
'Campbells borrowed code to escape the output 15/6/2006
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />")
Escape = strReturn
End Function

'*******declaring the variables****
  dim rsSearchH
  dim rsSearchM
  dim rsSearchL 
  dim rsFillFaculty
  dim rsFillLocation
  dim rsSearchFaculty
  dim Conn 
  dim strSQL
  dim strFacultyName
  dim strGivenName
  dim strSurname
  dim strName
  dim dtDate
  dim cboVal
  dim cboValDummy
  dim numOptionId
  dim numSupervisorId
  
  
       QORAtype = request.form("QORAtype")
	  session("QORAtype") = QORAtype
        session("searchType") = QORAtype
	  
	  numOperationID = request.form("cboOperation")
      session("cboOperation") = numOperationId
	  
    cboFacility = request.form("cboFacility")
	  session("cboFacility") = cboFacility
      
      numSupervisorId = session("numSupervisorId")

        strName = session("strName")
        strFacultyName = session("strFacultyName")
  '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
  
    

 if(QORAtype = "location") then 

 strSQL = "SELECT * FROM tblQORA, tblFacility,tblBuilding,tblCampus, tblRiskLevel ,tblFacilitySupervisor "_
  &" WHERE tblQORA.numFacilityId = tblFacility.numFacilityID and "_
  &" tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
  
  &" tblQORA.numFacilityId = "& cboFacility &" and "_
  &" tblFacility.numBuildingId = tblBuilding.numBuildingID and "_
  &" tblBuilding.numCampusId = tblCampus.numCampusID and "_
  
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "
    if not session("isAdmin") then
        strSQL = strSQL &" and tblFacilitySupervisor.numSupervisorId = "& numSupervisorID 
    end if
     strSQL = strSQL&" ORDER BY tblRiskLevel.numGrade, strTaskDescription"
 end if
 
 
 if(QORAtype = "operation") then
	 strSQL = "SELECT * FROM tblQORA, tblOperations, tblRiskLevel ,tblFacilitySupervisor "_
  &" WHERE tblQORA.numOperationId = tblOperations.numOperationId and "_
  &" tblFacilitySupervisor.numSupervisorID = tblOperations.numFacilitySupervisorID and"_
  
  &" tblQORA.numOperationId = "& numOperationID &" and "_
  &" tblQORA.strAssessRisk = tblRiskLevel.strRiskLevel "
  if not session("isAdmin") then
    strSQL = strSQL&" and tblFacilitySupervisor.numSupervisorId = "& numSupervisorID
  end if
    strSQL = strSQL&" ORDER BY tblRiskLevel.numGrade, strTaskDescription"
 end if
    'response.write (strSQL)
      
    set rsSearchH = server.CreateObject("ADODB.Recordset")
    rsSearchH.Open strSQL, Conn, 3, 3 
    
%>
	

<body>
    <!--#include file="HeaderMenu.asp" -->
    <%
        dim confirmationMsg
        confirmationMsg = request.form("confirmationMsg")
        if confirmationMsg <> "" then
         %>
            <div class="wrapper">
              <div class="content">
                <h1 class="pagetitle"><%=confirmationMsg%> </h1>
              </div>
            </div>
    <% end if %>

<div id="wrapper">
  <div id="content">
  <!-- Break out of frame --> 
      <%if rsSearchH.EOF <> true then  %>
  <form target="_blank" action="SupRDateModified-print.asp">
    <h1 class="pagetitle">Risk Assessment Search Results &nbsp;&nbsp;&nbsp;<input type="submit" value="Print preview" /></h1>    
  </form>
      <% end if %>
</div>

<% if rsSearchH.EOF <> true then  %>
    
    <% if(QORAtype="location") then %>
<table class="suprreportheader">
	<tr>
		<th>Campus:</th><td><%=cstr(rsSearchH("strCampusName")) %></td>
		<th>Building:</th><td><%=cstr(rsSearchH("strBuildingName")) %></td>
		<th>Room Name:</th><td><%=cstr(rsSearchH("strRoomName"))%></td>
		<th>Room Number:</th><td><%=cstr(rsSearchH("strRoomNumber"))%></td>
  
    </tr>
    <tr>
    	<th>Supervisor:</th><td><%=strName%></td>
    	<th>Faculty:</th><td colspan="5"><%=strFacultyName %></td>
 </tr>
 </table>
<% end if
if (QORAtype = "operation") then %>
<table class="suprreportheader">
	<tr>
		<th>Supervisor: </th><td><%=strName%></td>
		<th>Operation: </th><td><%=rsSearchH("strOperationName")%></td>	
	</tr>
</table>
<% end if
   'Response.Write(strSQL) 
  if not rsSearchH.EOF then 
       %>
    <br />
    <table class="sortable suprlevel" id="id13">
     <caption>
      To sort any column, click on a table heading. 
      </caption>
      <thead>
        <tr>
            <th style="width:130px">&nbsp;</th>
        	<th class="qoraID">Ra No.</th>
          	<th class="haztaskresult">Task</th>
    		<th class="assochazards">Hazards</th>
    		<th class="currentcontrols">Current Controls</th>
    		<th class="risklevel">Risk Level</th>
    		<th class="furtheraction">Proposed Controls</th>
    		<th class="renewaldate">Review Date</th>
    		<th class="swms">SWMS</th>
        </tr>
      </thead>
    	<tbody>
      <%

	 while not rsSearchH.EOF 
    dtDate = dateAdd("yyyy",2,rsSearchH("dtDateCreated"))
    
    %>
      
        <tr>
            <td><a href="EditQORA.asp?numCQORAId=<%=rsSearchH("numQORAID")%>">Edit</a> / 
                <a href="#" data-toggle="modal" data-target="#CopyModal" data-qora="<%=rsSearchH("numQORAID")%>">Copy</a> / 
                <a href="#" data-toggle="modal" data-target="#ArchiveModal" data-qora="<%=rsSearchH("numQORAID")%>">Archive</a>


            </td>
        <td><%=Escape(rsSearchH("numQORAId"))%></td>
          <!--td><a title="Click to edit this Risk Assessment." href="EditQORA.asp?numCQORAId=<%=rsSearchH("numQORAID")%>"><%=rsSearchH("strTaskDescription")%></a></td-->
		  <td><%=rsSearchH("strTaskDescription")%></td>
          <!--		<td><% Response.Write(rsSearchH(11))%></td> -->
          <td><%=Escape(rsSearchH("strHazardsDesc"))%></td>
          <td><%
          
          testval = rsSearchH("numQORAId")
           	'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& "<br/>"
     		' get the next record
           rsControls.MoveNext
     		wend 
     	   %>
     	  
     	<%=strControlsImplemented%>
          
       </td>
          <td><center>
              <%=rsSearchH("strAssessRisk")%>
            </center></td>
         <!-- old 'further action required' code <td><% Response.Write(rsSearchH("strText"))%>
            <%if rsSearchH("boolFurtherActionsSWMS")= true then %>
            <BR>
            <a target="_blank" href="http://www.ehs.uts.edu.au/forms/swms.doc" title="Safe Work Method Statement (in Microsoft Word format, 47 Kb).">Safe Work Method Statement</a>
            <%end if%>
            <%if rsSearchH("boolFurtherActionsChemicalRA")= true then %>
            <BR />
            <a target="_blank" href="http://www.ocid.uts.edu.au/" title="Chemical risk assessment at OCID.">Chemical Risk Assessment</a>
            <%end if%>
            <%if rsSearchH("boolFurtherActionsGeneralRA")= true then %>
            <BR />Detailed Risk Assessment<%end if%></td>
          <td><% Response.Write(rsSearchH(17))%></td>-->
          <td><%
          ' New code to put in the unimplemented risk controls
          
          testval = rsSearchH("numQORAId")
           	'here we need to populate the textarea with any existing controls we can locate
        	set connControls = Server.CreateObject("ADODB.Connection")
  			connControls.open constr
			' setting up the recordset
   			strControls ="Select * from tblRiskControls where numQORAId = "&testval&" and not boolImplemented"
  			set rsControls = Server.CreateObject("ADODB.Recordset")
        	rsControls.Open strControls, connControls, 3, 3
        	strControlsImplemented =""
        	while not rsControls.EOF 
         		strControlsImplemented = strControlsImplemented +rsControls("strControlMeasures")& "<br/>"
     		' get the next record
           rsControls.MoveNext
     		wend 
     	   %>
     	  
     	<%=strControlsImplemented%>
          
       </td>
          <%
              dim today
         today = Date()
               %>
      	<td <% If not isNull(rsSearchH("dtReview")) and DateDiff("d", rsSearchH("dtReview"), today) > 1 Then %>style="color:red;font-weight:bold" <%end if %> ><center><%=rsSearchH("dtReview")%></center></td>
         <td><center>
        <% If rsSearchH("boolSWMSRequired") = true Then %>
                 <form method="post" action="SWMS.asp">
         <input type="submit" value="SWMS" name="btnSWMS" />
         <input type="hidden" name="hdnQORAId" value="<%=rsSearchH("numQORAId")%>" />
         <input type="hidden" name="hdnNoSaveBeforeSWMS" value="nosave"/>
         </form>

        <% End if%>
                 </center></td>
            </tr>
        <%
    rsSearchH.MoveNext  
 wend

 %>
      </tbody>
    </table>
    <%
 'end if 
 end if %>
 
<%else%>
<p style="background-color:#fff">There are currently no Risk Assessments for this facility or operation</p>
<%end if%>
</div>


    <%
        
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
        
         %>

<div class="modal fade" id="CopyModal" tabindex="-1" role="dialog" aria-labelledby="CopyModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
          <h4 class="modal-title" id="exampleModalLabel">New message</h4>
      </div>
      <div class="modal-body">
        <form id="copyForm">
          <input type="hidden" class="form-control" id="qora" name="qora"/>
            <input type="hidden" name="mode" ID="QORAtype" value=""/>
          <div class="form-group">
            <label for="recipient-name" class="control-label">To Facility:</label>
           <select class="form-control" autocomplete="off" id="myfacility" size="1" name="cboFacility" tabindex="1" onchange="$('#QORAtype').val('location');$('#submitCopy').html('Copy to Location');">
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
          </div>
            <hr/>
            <b>OR</b>
            <hr />
           <div class="form-group">
            <label for="recipient-name" class="control-label">To Operation:</label>
           <select class="form-control" autocomplete="off" id="myoperation" name="cboOperation" id="cboOperation" Onchange="$('#QORAtype').val('operation');$('#submitCopy').html('Copy to Operation');">
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
          </div>
        </form>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
        <button type="button" id="submitCopy" class="btn btn-primary">Copy to..</button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="ArchiveModal" tabindex="-1" role="dialog" aria-labelledby="ArchiveModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
          <h4 class="modal-title" id="H1">New message</h4>
      </div>
      <div class="modal-body">
        <form id="archiveForm">
      
            <input type="hidden" name="mode" value="archive"/>
            <input type="hidden" name="superv" value="<% =session("numSupervisorId") %>" />
            <input type="hidden" name="qora" id="archiveQora" value=""/>
          <div class="form-group">
            <label  class="control-label">Are you sure you wish to archive this Risk Assessment?</label>
          </div>
        </form>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
        <button type="button" id="submitArchive" class="btn btn-primary">Archive</button>
      </div>
    </div>
  </div>
</div>

    <script type="text/javascript">
        $('#CopyModal').on('show.bs.modal', function (event) {
            var button = $(event.relatedTarget) // Button that triggered the modal
            var qora = button.data('qora') // Extract info from data-* attributes
            // If necessary, you could initiate an AJAX request here (and then do the updating in a callback).
            // Update the modal's content. We'll use jQuery here, but you could use a data binding library or other methods instead.
            var modal = $(this)
            modal.find('.modal-title').text('Copy Risk Assessment: ' + qora)
            modal.find('.modal-body #qora').val(qora)
        })

        $(function () {
            //twitter bootstrap script
            $("#submitCopy").click(function () {
                $.ajax({
                    type: "POST",
                    url: "AJAXCopy.asp",
                    data: $('#copyForm').serialize(),
                    success: function (data) {
                        var obj = jQuery.parseJSON(data);
                        var newRA = obj.result;
                        alert("Copied to RA "+newRA);
                        $("#CopyModal").modal('hide');
                        $("#refreshResults").submit();
                    },
                    error: function () {
                        alert("failure");
                    }
                });
            });
        });

        $('#ArchiveModal').on('show.bs.modal', function (event) {
            var button = $(event.relatedTarget) // Button that triggered the modal
            var qora = button.data('qora') // Extract info from data-* attributes
            // If necessary, you could initiate an AJAX request here (and then do the updating in a callback).
            // Update the modal's content. We'll use jQuery here, but you could use a data binding library or other methods instead.
            var modal = $(this)
            modal.find('.modal-title').text('Archive Risk Assessment: ' + qora)
            modal.find('.modal-body #archiveQora').val(qora)
        })

        $(function () {
            //twitter bootstrap script
            $("#submitArchive").click(function () {
                $.ajax({
                    type: "POST",
                    url: "AJAXArchive.asp",
                    data: $('#archiveForm').serialize(),
                    success: function (data) {
                        var obj = jQuery.parseJSON(data);
                        var newRA = obj.result;
                        alert("Archived RA " + newRA);
                        $("#ArchiveModal").modal('hide');
                        $("#refreshResults").submit();
                    },
                    error: function () {
                        alert("failure");
                    }
                });
            });
        });

    </script>

    <form id="refreshResults" action="SupRDateModified.asp" method="post">
        <input type="hidden" name="QORAtype" value="<%=QORAtype %>" />
        <input type="hidden" name="cboOperation" value="<%=numOperationID %>" />
        <input type="hidden" name="cboFacility" value="<%=cboFacility %>" />
    </form>
</body>
</html>
