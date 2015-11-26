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
Dim rsFillBuilding
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblOperations order by strOperationName"
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - Delete Facility</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ((document.DeleteFacility.cboBuildingName.value != "0") && (document.DeleteFacility.cboRoomNumber.value != "0") && (document.Form2.txtRoomName.value !="")) 
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

// function to reload the form to add the new entries
function FillOperation()
{
 document.DeleteOperation.submit();

}
// to reset the form
function resetForm()
{
	document.deletefacility.cboBuildingName.selectedIndex = 0
}
</script>
</head>

<body>
    <!--#include file="adminMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Delete a Facility</h1>
 
 <center>

<table class="adminfn" style="width: 55%">
<form method="post" action="DeleteOperation.asp" name="DeleteOperation">
<tr>
 <th>Existing Operation Name</th>
 <td>

  <%    numOperationID = cint(request.form("cboOperationName"))
        if numOperationID = "" then
	       numOperationID = 0
        end if %>
        <select size="1" name="cboOperationName" onChange="javascript:FillOperation()">
        <option value="0"
         <% if numOperationID = 0 then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <%while not rsFillBuilding.Eof%>
        <option value="<%=rsFillBuilding("numOperationID")%>"
        <% if rsFillBuilding("numOperationID") = numOperationID then
		  response.Write "selected='selected'"
		  end if %>><%=cstr(rsFillBuilding("strOperationName")) %></option>
        <%rsFillBuilding.Movenext
         wend 
         
         ' closing the connections
         
           rsFillBuilding.close
           set rsFillBuilding = nothing
           conn.Close
           set conn = nothing
         %>

        </select></td>
      </tr>
      
    </form>
    
    <form method="post" action="AdminDelete.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice(this);">
    <input type="hidden" name="hdnOperationId" value="<%=numOperationID%>" />  
    <input type="hidden" name="hdnOption" value="Operation" />  
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
 <th>Existing Operation Name</th>
 <td><input type="text" name="txtOperationName" size="35" value="<%=strOperationName %>"/></td>
</tr>
  
<tr>
 <td colspan="2">
 <center>
   <input type="submit" value="Delete" name="btnSave" />
 </center>
 </td>
</tr>
  </form>
</table>

</center>

</div></div>

</body>

</html>