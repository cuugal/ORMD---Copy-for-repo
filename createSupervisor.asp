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
  <!--#include file="bootstrap.inc"--> 
<title>Online Risk Register - Create a User</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{

  if ((document.Form1.strAccessLevel.value != 'S' || document.Form1.cboFaculty.value != "0") && (document.Form1.txtSurname.value != "") && (document.Form1.txtGivenName.value !="")
   && (document.Form1.txtLoginId.value !="") && (document.Form1.txtPassword.value !=""))
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
 
      alert ("Any fields on the form cant be empty, please fill in the entire form !");
    return(false);
 }
}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content" class="contentcenter">

 <h2 class="pagetitle">Create a User</h2>
 


 <form method="post" action="AdminCreate.asp" name="Form1" onsubmit="return ConfirmChoice();">
<table class="adminfn" >

    <tr>
    <th>User Type:</th>
    <td>

        <select name="strAccessLevel" id="strAccessLevel">
        <option value="S">Supervisor</option>
        <option value="T">Assessor</option>
        </select>

    </td>
    </tr>


  <tr id="userfaculty">
   <th>Existing Faculty/Unit:</th>
   <td>
     <select size="1" name="cboFaculty" tabindex="0">
      <option value="0">Select any one</option>
      <%while not rsfillFaculty.EOF %>
      <option value="<%=rsFillFaculty("numFacultyId")%>"><%=rsFillFaculty("strFacultyName")%></option>
      <%rsfillFaculty.Movenext
        wend%>
     </select></td>
    </tr>

 <tr>
  <th>New User Given Name:</th>
  <td><input type="text" name="txtGivenName" size="20" tabindex="1" /></td>
 </tr>
 
 <tr>
   <th>New User Surname:</th>
   <td><input type="text" name="txtSurname" size="20" tabindex="2" /></td>
 </tr>
 
 <tr>
   <th>New Login ID:</th>
   <td><input type="text" name="txtLoginId" size="20" tabindex="3" /></td>
 </tr>

 <tr>
  <th>New Password:</th>
  <td><input type="password" name="txtPassword" size="20" tabindex="4" /></td>
 </tr>

 <tr>

  <td colspan="2">

    <input type="hidden" name="hdnOption" value="Supervisor" />&nbsp;
    <input type="submit" value="Save" name="btnSave" tabindex="5" />
    <input type="reset" value="Clear" name="btnClear" tabindex="6" />

  </td>
 </tr>
 </table>
<script type="text/javascript">

$("#strAccessLevel").change(function(){

    if($(this).val() == 'S'){
        $("#userfaculty").show();

    }
    else{
        $("#userfaculty").hide();
    }
});
</script>
</form>



</div></div>

</body>
</html>