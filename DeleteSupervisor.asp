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

<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
	answer = confirm("Are you sure that you want to permanently delete this record from the database?")
		if (answer == true) 
		{ 
           return ;
		} 
		else
		 { 
		 return (false);
		}
}

// function to reload the form to add the new entries
function FillDetails()
{
 document.DeleteSupervisor.submit();

}
</script>
<%
Dim conn
Dim rsFillLoginId
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor order by strLoginId"
   set rsFillLoginId = Server.CreateObject("ADODB.Recordset")
   rsFillLoginId.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - Delete a Supervisor</title>
</head>

<body>

<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Delete a Supervisor</h1>
 
 <center>

<table class="adminfn" style="width: 55%">
 <tr>
   <th>Existing Login ID:</th>
   <td><form method="post" action="DeleteSupervisor.asp" name="DeleteSupervisor">
        <% strLoginID = request.form("cboLoginId")
        if strloginID = "" then
	       strLoginID = "0"
        end if %>
      <select size="1" name="cboLoginId" onchange="javascript:FillDetails()">
      <option value="0" 
        <% if strLoginID = "0" then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <% while not rsFillLoginId.Eof%>
        <option value="<%=rsFillLoginId("strLoginID")%>"
        <% if rsFillLoginId("strLoginID") = strLoginID then
		  response.Write "selected='selected'"
	  
		  end if %>><%=rsFillLoginId("strLoginID")%></option>
        <%rsFillLoginId.Movenext
         wend 
         
         ' closing the connections
         
           rsFillLoginId.close
           set rsFillLoginId = nothing
           conn.Close
           set conn = nothing
         %>
      </select></form>
	  </td>
</tr>
<tr>
  <th>Supervisor Surname</th>
  <form method="post" action="AdminDelete.asp" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">

<%
Dim connDet
Dim rsFillDetails


'Database Connectivity Code 
  set connDet = Server.CreateObject("ADODB.Connection")
  connDet.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor where strLoginId = '"& strLoginId &"'"
   set rsFillDetails = Server.CreateObject("ADODB.Recordset")
   rsFillDetails.Open strSQL, connDet, 3, 3
   
%>
 <td>
      <% if rsFillDetails.EOF <> true then%>
      <input type="text" name="txtSurname" size="20" tabindex="1" value="<%= rsFilldetails("strSurname")%>" /></td>
      
      <th>Supervisor's Faculty/Unit</th>
      <td>
     <%'code to add the related faculties for that str loginID 
      Dim connFac
      Dim rsFillFac
     
     'Database Connectivity Code 
      set connFac = Server.CreateObject("ADODB.Connection")
      connFac.open constr
 
     ' setting up the recordset
       strSQL ="SELECT * FROM tblFacilitySupervisor, tblFaculty WHERE tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyID and strLoginId = '"& strLoginId &"'"
       set rsFillFac = Server.CreateObject("ADODB.Recordset")
       rsFillFac.Open strSQL, connFac, 3, 3
      %>
        <%'code to add the different faculties for that str loginID 
      Dim connFaculty
      Dim rsFillFaculty
     
     'Database Connectivity Code 
      set connFaculty = Server.CreateObject("ADODB.Connection")
      connFaculty.open constr
 
     ' setting up the recordset
       strSQL ="Select * from tblFaculty order by strFacultyName"
       set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
       rsFillFaculty.Open strSQL, connFac, 3, 3
      %>
      <select size="1" name="cboFaculty" tabindex="3">
      <option value=<%=rsFillFac("tblFacilitySupervisor.numFacultyID")%> selected ><%=rsFillFac("strFacultyName")%></option>
      <%While not rsFillFaculty.EOF%>
      <option value="<%=rsFillFaculty("numFacultyID") %>"><%=rsFillFaculty("strFacultyName")%></option>
      <% rsFillFaculty.MoveNext
        wend %> 
      </select></td>
      
    </tr>
    
    <tr>
      <th>Supervisor's Given Name</th>
      <td>
       <input type="text" name="txtGivenName" size="20" tabindex="2" value="<%= rsFillDetails("strGivenName")%>" /></td>
    
      <th>Supervisor's Password:</th>
      <td>
       <input type="password" name="txtPassword" size="20" tabindex="4" value="<%= rsFillDetails("strPassword")%>" /></td>
    </tr>
    <tr>
      <td width="29%">&nbsp;</td>
      <td width="20%">
      <p align="center"><br>
    <%else%>
    <input type="text" name="txtSurname" size="20" tabindex="1" value="" /></td>
      
      <th>Supervisor Faculty/Unit</th>
      <td>
      <select size="1" name="cboFaculty" tabindex="3">
       <option>No Records</option>
      </select></td>
    </tr>
    
    <tr>
      <th>Supervisor's Given Name</th>
      <td>
      <input type="text" name="txtGivenName" size="20" tabindex="2" value="" /></td>
    
      <th>Supervisor's Password:</th>
      <td>
      <input type="password" name="txtPassword" size="20" tabindex="4" value="" /></td>
    </tr>
    <tr>
      <td width="29%">&nbsp;</td>
      <td width="20%">
      <p align="center"><br>

    <%end if%>     
   <input type="hidden" name="hdnLoginId" value="<%=strLoginId%>" />
   <input type="hidden" name="hdnOption" value="Supervisor" />
   <input type="submit" value="Delete" name="btnSave" tabindex="5" /></td>


    </tr>

  </form>
</table>

</center>

</div></div>

</body>

</html>