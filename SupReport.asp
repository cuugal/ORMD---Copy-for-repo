
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>


<script type="text/javascript">
// function to ask about the confirmation of the file. ADDED BY DLJ 7June7
function ConfirmChoice() 
{ 
  if ((document.FormA.cboFacility.value == "0") && (document.FormA.cboOperation.value == "0") ) 
  {
      alert ("Please select a Facility or Operation to proceed.");
	   return(false);
	}
}
function ChangeType(val)
{
 document.FormA.QORAtype.value = val;
 //console.log(document.Form2.QORAtype.value);

}
// function to reload the form to add the new entries
function FillDetails()
{
 document.Form1.submit();

}

</script>

<%dim hTaskDesc
  dim hNavigCnt
  
  hTaskDesc = Request.Form("hdnTaskDesc") 
  hNavigCnt = Request.Form("hdnNavigationCnt")
  if hNavigCnt = "1" then
  %>
  The &quot;<%=hTaskDesc%>&quot; Risk Assessment has been successfully added to the Online Risk Register.
  <%

  end if
  
  dim pn 
  dim htd
  
  pn = session("pn")
  htd = session("HTask")
  
  if pn = "1" then
    %>
  The &quot;<%=htd%>&quot; RA has been successfully edited.
  <%
  
  'Clear the session variables so the message won't persist
  session("pn") = ""
  session("HTask") = ""
  end if
  
%>
<%
 dim rsSearchFacility
 dim rsSearchFaculty
 dim Conn
 dim strSQL
 
   '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
        
  '------------------------get the facilities for the login ---------------
  'AA jan 2010 as part of reln fix join to tblFacilitySupervisor
  strSQL = "Select * "_
  &" from tblfacility "_
  &" where tblFacility.numFacilitySupervisorId = "& session("numSupervisorId") &"  order by strRoomName" 

  set rsSearchFacility = server.CreateObject("ADODB.Recordset")
  rsSearchFacility.Open strSQL, Conn, 3, 3 
    

  'AA jan 2010 add join on tblFacilitySupervisor as part of reln repair
  strSQL = "Select * from tblOperations where numFacilitySupervisorId = "&session("numSupervisorId")

  set rsOper = server.CreateObject("ADODB.Recordset")
  rsOper.Open strSQL, Conn, 3, 3
  %>


<body>

<div id="wrapper">
<div id="content">

<h1 class="pagetitle">My Risk Assessments</h1>

<center>

<table class="myqora" style="width: 46%; ">
<thead>
<tr>
 <th colspan="3">My Risk Assessments</th>
</tr>
</thead>
<tbody>
<tr>
<!-- <form method="post" action="SupRDateModified.asp">  -->
<form method="post" action="SupRDateModified.asp"  name="FormA" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
<input type="hidden" name="QORAtype" value=""/>

    <th>Faculty/Unit</th>
    <td colspan="2"><%response.write(strFacultyName) %></td>
</tr>
<tr>
    <th>Supervisor Name</th>
    <td colspan="2"><%response.write(strName) %></td>
</tr>
<tr>
    <th>Select Facility</th>
    <td>
      <select size="1" name="cboFacility" onchange="javascript:ChangeType('location')">
      <option value="0">Select Facility</option>
      <%while not rsSearchFacility.EOF    %>
      <option value="<%=rsSearchFacility("numFacilityId")%>"><%=cstr(rsSearchFacility("strRoomNumber"))+" / "+cstr(rsSearchFacility("strRoomName")) %></option>
      <%rsSearchFacility.MoveNext 
      wend  %>
      </select>&nbsp;
      </td><td>
      <input type="submit" size="70" value="Generate Facility Report" name="btnGenRep" /></td>
      
      </td>
    </tr>

    <tr>
    <th>Select Operation/Project</th>
    <td>
      <select size="1" name="cboOperation" onchange="javascript:ChangeType('operation')">
      <option value="0">Select Operation/Project</option>
      <%while not rsOper.EOF    %>
      <option value="<%=rsOper("numOperationId")%>"><%=cstr(rsOper("strOperationName")) %></option>
      <%rsOper.MoveNext 
      wend  %>
      </select>&nbsp;
      </td><td>
      <input type="submit" size="70"value="Generate Operation/Project Report" name="btnGenRep" /></td>
      </form>
      </td>
    </tr>
    </tbody>
  </table>
</div>
</div>

</body>
