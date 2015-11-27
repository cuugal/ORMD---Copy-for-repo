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
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.Form1.txtCampusName.value !="") 
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
<title>Online Risk Register - Create a Campus</title>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
<div id="content">

<h1 class="pagetitle">Create a Campus</h1>

<center>

<form method="post" action="AdminCreate.asp" name="Form1" onsubmit="return ConfirmChoice();">

<table class="adminfn" style="width: 55%">
  <tr>
   <th>New Campus Name</th>
   <td><input type="text" name="txtCampusName" size="35" tabindex="0" /></td>
 </tr>
 <tr>
   <td colspan="2">
    <center>
    <input type="submit" value="Save" name="btnSave" tabindex="1" />
    <input type="reset" value="Clear" name="btnClear" tabindex="2" />
    <input type="hidden" name="hdnOption" value="Campus" />
    </center>
   </td>
 </tr>
</table>



</form>

</center>
</div><!-- close #content -->
</div><!-- close #wrapper -->

</body>
</html>