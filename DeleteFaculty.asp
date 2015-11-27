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
Dim rsFillFaculty
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFaculty order by strFacultyName"
   set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
   rsFillFaculty.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
 <title>Online Risk Register - Delete a Faculty/Unit</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ( document.Form1.cboFacultyName.value !="0") 
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
	
      alert ("Any fields on the form cant be empty , please fill in the entire form !");
	   return(false);
	}
}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Delete a Faculty/Unit</h1>
 
 <center>

<form method="post" action="AdminDelete.asp" name="Form1" onsubmit="return ConfirmChoice();">
<table class="adminfn" style="width: 65%">
<tr>
 <th>Existing Faculty/Unit:</th>
 <td>
  <select size="1" name="cboFacultyName">
  <option value="0">Select any one</option>

  <%While not rsFillFaculty.EOF %>
  <option value="<%=rsFillFaculty("numFacultyId")%>"><%=rsFillFaculty("strFacultyName")%></option>
  <%
  rsFillFaculty.Movenext
  wend%>
  </select>
  </td>
</tr>

<tr>
 <td colspan="2">
 <center>
   <input type="submit" value="Delete" name="btnSave" />&nbsp;<input type="hidden" name="hdnOption" value="Faculty" />
 </center>
 </td>
</tr>
</table>

</form>

</center>

</div></div>

</body>

</html>