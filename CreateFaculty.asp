<%@Language = VBscript%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<%
 if session("strLoginId") <> "admin" then
  response.redirect "AccessRestricted.htm"
 end if
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
<title>Online Risk Register - Create a Faculty/Unit</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.Form1.txtFacultyName.value !="" && document.Form1.txtDGivenName.value !="" && document.Form1.txtDSurName.value !="" && document.Form1.txtDLoginId.value !="" && document.Form1.txtDPassword.value !="" ) 
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
 
      alert ("No field on this form can be empty - please complete the entire form.");
    return(false);
 }
}
</script>
</head>

<body>
    <!--#include file="adminMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Create a Faculty/Unit</h1>
 
 <center>

 <form method="post" action="AdminCreate.asp" name="Form1" onsubmit="return ConfirmChoice();" >

 <table class="adminfn" style="width: 55%;">
  <tr>
   <th>New Faculty/Unit Name:</th>
   <td><input type="text" name="txtFacultyName" size="35" tabindex="0" /></td>
 </tr>
 <tr>
   <th>Dean/Director's Given Name:</th>
   <td><input type="text" name="txtDGivenName" size="35" tabindex="1" /></td>
 </tr>
 <tr>
   <th>Dean/Director's Surname:</th>
   <td><input type="text" name="txtDSurname" size="35" tabindex="2" /></td>
 </tr>
 <tr>
   <th>Dean/Director's Login ID:</th>
   <td><input type="text" name="txtDLoginID" size="20" tabindex="3" /></td>
 </tr>
 <tr>
   <th>Dean/Director's Password:</th>
   <td><input type="password" name="txtDPassword" size="20" tabindex="4" /></td>
 </tr>
 <tr>
   <td colspan="2">
    <center>
    <input type="submit" value="Save" name="btnSave" tabindex="5" />&nbsp;
    <input type="reset" value="Clear" name="btnClear" tabindex="6" />
    <input type="hidden" name="hdnOption" value="Faculty" />
    </center>
   </td>
 </tr>
 </table>

</form>

</center>

</div></div>

</body>
</html>