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
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
<title>Online Risk Register - Edit a Faculty/Unit</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.Form1.txtFacultyName.value !="" && document.EditFaculty.cboFacultyName.value !="0" && document.Form1.txtDGivenName.value !="" && document.Form1.txtDSurName.value !="" && document.Form1.txtDLoginId.value !="" && document.Form1.txtDPassword.value !="") 
  {
     answer = confirm("Do you want to save changes to this record to the database?")
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

// function to reload the form to add the new entries
function FillDetails()
{
 document.EditFaculty.submit();
}
</script>
<script language="vbscript" type="text/VBScript">
  function clearform()
  dim val 
  val = ""
  document.EditFaculty.cboFacultyName.value = 0
  document.Form1.txtFacultyName.value = val
  document.Form1.txtDGivenName.value = val
  document.Form1.txtDSurName.value = val
  document.Form1.txtDLogin.value = val
  document.Form1.txtDPassword.value = val
  end function
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Edit a Faculty/Unit</h1>
 
 <center>



 <table class="adminfn" style="width: 65%">
 <tr>
  <th>Existing Faculty/Unit:</th>
  <td>
  
  <form method="post" action="EditFaculty.asp" name="EditFaculty">
 <% numFacultyID = cint(request.form("cboFacultyName"))
        if numFacultyID = "" then
        numFacultyID = 0
        end if %>
        <select size="1" name="cboFacultyName" onchange="javascript:FillDetails()">
          <option value="0"
          <% if numFacultyId = "0" then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillFaculty.Eof%>
        <option value="<%=rsFillFaculty("numFacultyId")%>"
        <% if rsFillFaculty("numFacultyId") = numFacultyID  then
    response.Write "selected"
    end if %>><%=rsFillFaculty("strFacultyName")%></option>
        <%rsFillFaculty.Movenext
         wend %>
        </select></td>
       </form>
  </td>
 </tr>

 <tr>
  <th>New Faculty/Unit</th>
  <td>
   <form method="post" action="AdminEdit.asp" name="Form1" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
   <input type="hidden" name="hdnFacultyId" value="<%=numFacultyID%>" />
<% strSQL ="Select * from tblFaculty where numFacultyId = "& numFacultyID
      set rsFillF = Server.CreateObject("ADODB.Recordset")
      rsFillF.Open strSQL, conn, 3, 3
     
     if rsFillF.EOF <> True then
%>
     <input type="text" name="txtFacultyName" size="60" value="<%=rsFillF(1)%>"></td>
 </tr>

<tr>
 <th>New Dean/Director's Given Name:</th>
 <td><input type="text" name="txtDGivenName" size="35" value="<%=rsFillF(4)%>" /></td>
</tr>

<tr>
 <th>New Dean/Director's Surname:</th>
 <td><input type="text" name="txtDSurname" size="35" value="<%=rsFillF(5)%>" /></td>
</tr>

<tr>
 <th>New Dean/Director's Login ID:</th>
 <td><input type="text" name="txtDLogin" size="35" value="<%=rsFillF(2)%>" /></td>
</tr>

<tr>
 <th>New Dean/Director's Password:</th>
 <td><input type="password" name="txtDPassword" size="20" value="<%=rsFillF(3)%>" /></td>
</tr>

<tr>
 <td colspan="2">
 <center>
  <input type="submit" value="Save" name="btnSave" />&nbsp;
  
  
  <!-- CL note: This button does not work in Mozilla etc. as it calls VBscript - replaced with a reset button instead 
  <input type="button" value="Clear" name="btnClear" onclick="clearform()" />
  -->
  <input type="reset" value="Reset" name="btnClear" />

  <input type="hidden" name="hdnOption" value="Faculty" />
</center>
</td>

<%else%>

 <input type="text" name="txtFacultyName" size="60" value=""></td>
<%end if%>     
</tr>
 </table>

</form>

</center>

</div></div>

</body>
</html>