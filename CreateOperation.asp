<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
 if session("strLoginId") <> "admin" then
  response.redirect "AccessRestricted.htm"
 end if
%>
<%
Dim conn
Dim rsFillBuilding
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblBuilding,tblCampus where tblBuilding.numCampusId = tblCampus.numCampusId order by strBuildingName"
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
<title>Online Risk Register - Create an Operation</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ((document.Form1.cboSupervisorName.value != "0") && (document.Form1.txtOperationName.value !="")) 
  {
     answer = confirm("Do you want to save this record to the database?")
  if (answer == true) 
  { 
           return ;
  } 
  else
   { 
   return (false);
  }
    }
  else
 {
 
      alert ("Any fields on the form cant be empty , please fill in the entire form !");
    return(false);
 }
}
</script>
</head>

<body>

<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Create an Operation</h1>
 
 <center>

 <form method="post" action="AdminCreate.asp" name="Form1" onsubmit="return ConfirmChoice();">

<table class="adminfn" style="width: 55%">


<tr>
 <th>New Operation Name</th>
 <td><input type="text" name="txtOperationName" size="35" /></td>
</tr>
<tr>
 <th>Existing Supervisor</th>
 <td><%
Dim connSupervisor
Dim rsFillSupervisor

'Database Connectivity Code 
  set connSupervisor = Server.CreateObject("ADODB.Connection")
  connSupervisor.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor order by strGivenName"
   set rsFillSupervisor = Server.CreateObject("ADODB.Recordset")
   rsFillSupervisor.Open strSQL, connSupervisor, 3, 3
%>
        <select size="1" name="cboSupervisorName">
        <option value="0">Select any one</option>

        <%While not rsFillSupervisor.EOF %>
         <%if rsFillSupervisor("strAccessLevel") = "S" then%>
         <option value="<%=rsFillSupervisor("numSupervisorID")%>"><%=cstr(rsFillSupervisor("strGivenName")) + " " +cstr(rsFillSupervisor("strSurname"))%></option>
          <%
          end if
           rsFillSupervisor.Movenext
           wend%>

        </select>
      </td>
      </tr>
<tr>
 <td colspan="2">
 <center>
  <input type="submit" value="Save" name="btnSave" />&nbsp;<input type="reset" value="Clear" name="btnClear" />
  <input type="hidden" name="hdnOption" value="Operation" />
 </center>
 </td>
</tr>
 </table>

</form>

</center>

</div></div>

</body>
</html>