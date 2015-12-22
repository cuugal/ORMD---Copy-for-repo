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
Dim rsFillCampus
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblCampus order by strCampusName"
   set rsFillCampus = Server.CreateObject("ADODB.Recordset")
   rsFillCampus.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!--#include file="bootstrap.inc"--> 
 <title>Online Risk Register - Delete a Campus</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.Form1.cboCampusName.value != "0") 
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
		else
	{
	
      alert ("Please complete a fields in the form.");
	   return(false);
	}
}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h2 class="pagetitle">Delete a Campus</h2>
 
 <center>

<form method="post" action="AdminDelete.asp" name="Form1" onsubmit="return ConfirmChoice();">

<table class="adminfn" style="width: 45%">
<tr>
 <th>Existing Campus Name:</th>
 <td>
<select size="1" name="cboCampusName">
<option value="0">Select any one</option>

<%While not rsFillCampus.EOF %>
   <option value="<%=rsFillCampus("numCampusId")%>"><%=rsFillCampus("strCampusName")%></option>
   <%
   rsFillCampus.Movenext
   wend%>
  </select>
 </td>
</tr>
<tr>
 <td colspan="2">
 <center>
  <input type="submit" value="Delete" name="btnSave" />&nbsp;<input type="hidden" name="hdnOption" value="Campus" />
 </center>
 </td>
</tr>
</table>

</form>

</center>

</div></div>

</body>
</html>