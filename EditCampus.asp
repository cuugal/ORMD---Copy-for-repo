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
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
<title>Online Risk Register - Edit a Campus</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.EditCampus.cboCampusName.value != "0" && document.Form1.txtCampusName.value !="") 
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
 document.EditCampus.submit();
}

//  function clearform()
//{
//  alert("Exception Caught!");
//  document.EditCampus.cboCampusName.value == 1;
  //document.Form1.txtCampusName.value==" ";
// } 
 
</script>

<!-- note that VBscript doesn't work in Mozill etc. 
<script Language ="VBscript" type = "text/VBScript" >
  function clearform()
  dim val 
  val = ""
  document.EditCampus.cboCampusName.value = 0
  document.Form1.txtCampusName.value  = val
  end function
</script> -->
</head>

<body>
    <!--#include file="adminMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Edit a Campus</h1>
 
 <center>


 <table class="adminfn" style="width: 65%;">
  <tr>
   <th>Existing Campus Name</th>
   <td><form method="post" action="EditCampus.asp" name="EditCampus">
       <% numCampusID = cint(request.form("cboCampusName"))
        if numCampusID = "" then
        numCampusID = 0
        end if %>
        <select size="1" name="cboCampusName" onchange="javascript:FillDetails()">
        <option value="0"
          <% if numCampusId = "0" then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillCampus.Eof%>
        <option value="<%=rsFillCampus("numCampusId")%>"
        <% if rsFillCampus("numCampusId") = numCampusID  then
    response.Write "selected"
    end if %>><%=rsFillCampus("strCampusName")%></option>
        <%rsFillCampus.Movenext
         wend %>
         
        </select></form></td>
       </tr>

 <tr>
  <th>New Campus Name</th>

<form method="post" action="AdminEdit.asp" name="Form1" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
<input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" />

<% 
    strSQL ="Select strCampusName from tblCampus where numCampusId = "& numCampusID
   set rsFillC = Server.CreateObject("ADODB.Recordset")
   rsFillC.Open strSQL, conn, 3, 3

   if rsFillC.EOF <> True then
%>
 <td>
  <input type="text" name="txtCampusName" size="35" value="<%=rsFillc(0)%>">
 </td>

<%else%>

 <td>
  <input type="text" name="txtCampusName" size="35" value="" />
 </td>

<%end if%>
</tr>

 <tr>
   <td colspan="2">
   <center>
   <input type="submit" value="Edit" name="btnSave" />&nbsp;
  <!-- CL note: This button does not work in Mozilla etc. as it calls VBscript - replaced with a reset button instead 
  <input type="button" value="Clear" name="btnClear" onclick =clearform();>
  -->
  <input type="reset" value="Reset" name="btnClear" /><input type="hidden" name="hdnOption" value="Campus" /></td>
  </center>
  </td>
  </tr>
  

   
</tr>
</table>
</center>
</div>
  <p align="left">&nbsp;</p>
</form>

</body>

</html>