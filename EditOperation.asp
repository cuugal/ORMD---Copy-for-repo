<%@Language = VBscript%>
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - Edit a Facility</title>
 <script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
 if ((document.Form1.cboSupervisorName.value != "0") && (document.Form1.txtOperationName.value !="")) 
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
 
      alert ("No fields on this form may be empty - please fill in the entire form.");
    return(false);
 }
}
// function to reload the form to add the new entries
function FillOperation()
{
 document.EditOperation.submit();

}
</script>
<script type="text/vbscript">
  function clearform()
  dim val 
  val = " "
  document.EditFacility.cboBuildingName.value = 0
  document.EditFacility.cboRoomNumber.value = 0
  document.Form2.txtRoomNumber.value = val 
  document.Form2.txtRoomName.value = val 
  document.Form2.cboSupervisorName.value = 0

  end function
</script>
</head>
<%
 if session("strLoginId") <> "admin" then
  response.redirect "AccessRestricted.htm"
 end if
%>

<%
Dim conn
Dim rsFillOperation
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
   
   strSQL = "Select numOperationId, strOperationName from tblOperations order by strOperationName"
   set rsFillOperation = Server.CreateObject("ADODB.Recordset")
   rsFillOperation.Open strSQL, conn, 3, 3
%>
<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">
 <h1 class="pagetitle">Edit an Operation</h1>
 <center>

 <table class="searchtable" style="width: 65%;">
 <tr>
  <th>Existing Operation Name</th>
  <td>
   <form method="post" action="EditOperation.asp" name="EditOperation">
   <%    numOperationID = cint(request.form("cboOperationName"))
        if numOperationID = "" then
        numOperationID = 0
        end if %>
        <select size="1" name="cboOperationName" onchange="FillOperation()">
        <option value="0"
         <% if numOperationID = 0 then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillOperation.Eof%>
        <option value="<%=rsFillOperation("NumOperationID")%>"
        <% if rsFillOperation("NumOperationID") = numOperationID then
    response.Write "selected"
    end if %>><%=cstr(rsFillOperation("strOperationName"))%></option>
        <%rsFillOperation.Movenext
         wend 
         
         ' closing the connections
         
           rsFillOperation.close
           set rsFillOperation = nothing
           conn.Close
           set conn = nothing
         %>
        </select></td>
 </tr>
</form>
</td>
 </tr>
 
 <tr>
  <th><form method="post" action="AdminEdit.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
 <input type="hidden" name="hdnOperationId" value="<%=numOperationId%>" />
    <input type="hidden" name="hdnOption" value="Operation" />
	</th>
  <tr>
<%
Dim connOperationName
Dim rsFillOperationName

'Database Connectivity Code 
  set connOperationName = Server.CreateObject("ADODB.Connection")
  connOperationName.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblOperations where numOperationId = "& numOperationID 
   set rsFillOperationName = Server.CreateObject("ADODB.Recordset")
   rsFillOperationName.Open strSQL, connOperationName, 3, 3

Dim strOperationName
Dim strSupID

strOperationName = ""
strSupId = 0
if  rsFillOperationName.EOF <> True then   
   strOperationName = rsFillOperationName("strOperationName")
   strSupId = rsFillOperationName("numFacilitySupervisorID") 
end if
 %>
 <th>New Operation Name</th>
 <td><input type="text" name="txtOperationName" size="35" value="<%=strOperationName %>"/></td>
</tr>
<tr>
<%
Dim connSup
Dim rsFillSup

'Database Connectivity Code 
  set connSup = Server.CreateObject("ADODB.Connection")
  connSup.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor" 
   set rsFillSup = Server.CreateObject("ADODB.Recordset")
   rsFillSup.Open strSQL, connSup, 3, 3
%>
    <tr>
        <th>Supervisor Name</th>
        <td> 
        <!--<select size="1" name="cboSupervisorName">-->
        <select size="1" name="cboSupervisorID">
        <option value = "0">Select any one</option>
         <%while not rsFillSup.EOF %>
          <%if rsFillSup("strAccessLevel") = "S" then

             if (strSupID = rsFillSup("numSupervisorID")) then %>
        <option value = "<%=rsFillSup("numSupervisorID")%>" selected><%= cstr(rsFillsup("strGivenName")) + " " + cstr(rsFillsup("strSurName")) %></Option> 
       <% else%>
        <option value = "<%=rsFillSup("numSupervisorID")%>"><%= cstr(rsFillsup("strGivenName")) + " " + cstr(rsFillsup("strSurName")) %></Option> 
       <% end if %>
        <%
          end if 
          rsFillSup.MoveNext
          wend
        %></select></td>
      </tr>
      <tr>
        <td colspan="2"><center><input type="submit" value="Edit" name="btnSave" />&nbsp;
		<input type="reset" value="Reset this record" name="btnClear" />
		</form>
		</center>
		</td>
      </tr>
  </table>
  </center>
</div>
</div>
</body>
</html>