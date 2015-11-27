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
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - Edit a Building</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.EditBuilding.cboCampusName.value != "0" && document.EditBuilding.value != "0" && document.Form2.txtBuildingName.value !="") 
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
function FillBuildingCampus()
{
 document.EditBuilding.submit();
}
</script>
<script type="text/VBScript">
  function clearform()
  dim val 
  val = ""
  document.EditBuilding.cboCampusName.value = 0
  document.EditBuilding.cboBuildingName.value = 0
  document.Form2.txtBuildingName.value = val 
  end function
</script>
</head>

<body>
     <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Edit a Building</h1>
 
 <center>

<form method="post" action="EditBuilding.asp" name="EditBuilding" />

 <table class="adminfn" style="width: 65%">
 <tr>
  <th>
  <%
' Code to fill the campuses for editing the building

Dim connCampus
Dim rsFillCampus
Dim strSQL
dim numCampusID
dim numBuildingID

'Database Connectivity Code 
  set connCampus = Server.CreateObject("ADODB.Connection")
  connCampus.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblCampus order by strCampusName"
   set rsFillCampus = Server.CreateObject("ADODB.Recordset")
   rsFillCampus.Open strSQL, connCampus, 3, 3
   
   numCampusID = cint(request.form("cboCampusName"))
      if numCampusID = "" then
      numCampusID = 0
      end if
%>
Existing Campus Name</th>
<td>
 <select size="1" name="cboCampusName" onchange="javascript:FillBuildingCampus()">
 <option value="0" 
         <% if numCampusID = 0 then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillCampus.Eof%>
        <option value="<%=rsFillCampus("numCampusID")%>"
        <% if rsFillCampus("numCampusID") = numCampusID then
    response.Write "selected='selected'"
    end if %>><%=rsFillCampus("strCampusName")%></option>
        <%rsFillCampus.Movenext
         wend 
         
         ' closing the connections
         
           rsFillCampus.close
           set rsFillCampus = nothing
           connCampus.Close
           set connCampus = nothing
         %>
      </select></td>
    </tr>
     <tr>         
      <th>Existing Building Name</th>
      <%
' Code to fill the Building names for a particular campus

Dim connBuilding
Dim rsFillBuilding
 ' setting up the database connectivity
 
   set connBuilding = Server.CreateObject("ADODB.Connection")
   connBuilding.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblBuilding where numCampusId ="& numCampusID &" order by strBuildingname"
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, connBuilding, 3, 3
  
  numBuildingID = cint(request.form("cboBuildingName"))
    if numBuildingID = "" then
    numBuildingID = 0
    end if

%>
      <td>
      <select size="1" name="cboBuildingName" onchange="javascript:FillBuildingCampus()">
        <option value="0" 
         <% if numBuildingID = 0 then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillBuilding.Eof%>
        <option value="<%=rsFillBuilding("numBuildingID")%>"
        <% if rsFillBuilding("numBuildingID") = numBuildingID then
    response.Write "selected='selected'"
    end if %>><%=rsFillBuilding("strBuildingName")%></option>
        <%rsFillBuilding.Movenext
         wend 
     
          %>
      </select></form></td>
    </tr>
    <tr>
      <th>New Building Name</th>
      <td>
      <form method="post" action="AdminEdit.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
      <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" />
      <input type="hidden" name="hdnBuildingId" value="<%=numBuildingID%>" />
      <input type="hidden" name="hdnOption" value="Building" />
      <%        
            
             strSQL ="Select strBuildingName from tblBuilding where numBuildingId = "& numBuildingID
      set rsFillB = Server.CreateObject("ADODB.Recordset")
      rsFillB.Open strSQL, connBuilding, 3, 3
     
     if rsFillB.EOF <> True then
    %>
        <input type="text" name="txtBuildingName" size="35" value="<%=rsFillB(0) %>" /></td>
    <%else%>
            <input type="text" name="txtBuildingName" size="35" value="" /></td>
        <%end if%>  
    </tr>

   <tr>
    <td colspan="2">
    <center>
  <input type="submit" value="Save" name="btnSave" />&nbsp;
  <!-- CL note: This button does not work in Mozilla etc. as it calls VBscript - replaced with a reset button instead 
  <input type="Button" value="Clear" name="btnClear" onclick = clearform()>
  -->
  <input type="reset" value="Reset Building Name" name="btnClear" /></form>
  </center>
  </td>
 </tr>
 </table>

</form>

</center>

</div></div>

</body>
</html>