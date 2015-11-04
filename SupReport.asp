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
<%dim loginId
loginId = session("strLoginId")%>

<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
 <title>Online Risk Register - My Risk Assessments</title>

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
 dim rsSearchFacility1
 dim rsSearchFacility2
 dim rsSearchFaculty
 dim Conn
 dim strSQL
 
   '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
        
  '------------------------get the facilities for the login ---------------
  'AA jan 2010 as par tof reln fix join to tblFacilitySupervisor
  strSQL = "Select * "_
  &" from tblfacility, tblFacilitySupervisor"_
  &" where tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
  &" and tblFacilitySupervisor.strLoginID = '"& loginId &"'  order by strRoomName" 
  
  set rsSearchFacility1 = server.CreateObject("ADODB.Recordset")
  rsSearchFacility1.Open strSQL, Conn, 3, 3   
 
 'Get the login ID out as this is more useful
  strSQL = "Select * from tblFacilitySupervisor where strLoginID = '"& loginId &"'" 
  
  set rsID = server.CreateObject("ADODB.Recordset")
  rsID.Open strSQL, Conn, 3, 3 
  numSupervisorId = rsID("numSupervisorId")
 %>
 <%  '------------------------get the faculty for the login ---------------
  strSQL = "Select * "_
  &" from tblfacilitySupervisor,tblFaculty "_
  &" where tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId "_
  &" and tblFacilitySupervisor.strLoginId = '"& loginId &"'" 
  
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  'Response.Write(strSQL) 
  rsSearchFaculty.Open strSQL, Conn, 3, 3     
  strFacultyName = rsSearchFaculty("strFacultyName")     
  strGivenName = rsSearchFaculty("strGivenName")
  strSurname = rsSearchFaculty("strSurname")
  strName = cstr(strGivenName) + " " + cstr(strSurname)
  
  %>
</head>

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

  <%
  'AA jan 2010 add join on tblFacilitySupervisor as part of reln repair
  strSQL = "Select * "_
  &" from tblfacility, tblFacilitySupervisor"_
  &" where tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
  &" and tblFacilitySupervisor.strLoginID = '"& loginId &"' order by strRoomName" 
  
  set rsSearchFacility2 = server.CreateObject("ADODB.Recordset")
  rsSearchFacility2.Open strSQL, Conn, 3, 3 %>

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
      <%while not rsSearchFacility2.EOF    %>
      <option value="<%=rsSearchFacility2(0)%>"><%=cstr(rsSearchFacility2(1))+" / "+cstr(rsSearchFacility2(2)) %></option>
      <%rsSearchFacility2.MoveNext 
      wend  %>
      </select>&nbsp;
      </td><td>
      <input type="submit" size="70" value="Generate Facility Report" name="btnGenRep" /></td>
      
      </td>
    </tr>
    
     <%
  'AA jan 2010 add join on tblFacilitySupervisor as part of reln repair
  strSQL = "Select * from tblOperations where numFacilitySupervisorId = "&numSupervisorId
  
  'response.write strSQL
  'response.end
  set rsOper = server.CreateObject("ADODB.Recordset")
  rsOper.Open strSQL, Conn, 3, 3 %>
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

</html>