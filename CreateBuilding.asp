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
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
 <title>Online Risk Register - Create a Building</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.Form1.cboCampusName.value != "0" &&  document.Form1.txtBuildingName.value !="") 
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

 <h1 class="pagetitle">Create a Building</h1>
 
 <center>

<form method="post" action="AdminCreate.asp" name="Form1" onsubmit="return ConfirmChoice();">
<table class="adminfn" style="width: 55%">
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
 </select></td>
</tr>
  <tr>
   <th>New Building Name:</th>
   <td><input type="text" name="txtBuildingName" size="35" /></td>
 </tr>
 <tr>
   <td colspan="2">
    <center>
     <input type="submit" value="Save" name="btnSave" />&nbsp;<input type="reset" value="Clear" name="btnClear" /><input type="hidden" name="hdnOption" value="Building" />
    </center>
   </td>
 </tr>
</table>

</form>

</center>

</div></div>

</body>

</html>