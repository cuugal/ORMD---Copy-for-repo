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
 <!--#include file="bootstrap.inc"--> 
 <title>Online Risk Register - Edit a Facility</title>
 <script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ((document.EditFacility.cboBuildingName.value != "0") && (document.EditFacility.cboRoomNumber.value != "0") && (document.Form2.cboSupervisorName.value != "0") && (document.Form2.txtRoomName.value !="") && (document.Form2.txtRoomNumber.value !="")) 
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
function FillBuildingCampus()
{
 document.EditFacility.submit();

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
Dim rsFillBuilding
Dim strSQL

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select distinct(tblFacility.numBuildingId)as NumBuildingID,tblCampus.strCampusName,tblBuilding.strBuildingName from tblBuilding,tblCampus,tblFacility where tblBuilding.numCampusId = tblCampus.numCampusId and tblFacility.numBuildingId = tblBuilding.numBuildingId order by strBuildingName"
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, conn, 3, 3
%>
<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">
 <h2 class="pagetitle">Edit a Facility</h2>
 <center>

 <table class="adminfn" style="width: 65%;">
 <tr>
  <th>Existing Building Name</th>
  <td>
   <form method="post" action="EditFacility.asp" name="EditFacility">
   <%    numBuildingID = cint(request.form("cboBuildingName"))
        if numBuildingID = "" then
        numBuildingID = 0
        end if %>
        <select size="1" name="cboBuildingName" onchange="javascript:FillBuildingCampus()">
        <option value="0"
         <% if numBuildingID = 0 then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillBuilding.Eof%>
        <option value="<%=rsFillBuilding("NumBuildingID")%>"
        <% if rsFillBuilding("NumBuildingID") = numBuildingID then
    response.Write "selected"
    end if %>><%=cstr(rsFillBuilding("strBuildingName")) + " - " + cstr(rsFillBuilding("strCampusName")) + "  " + "Campus"%></option>
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
<%
Dim connRoom
Dim rsFillRoom

'Database Connectivity Code 
  set connRoom = Server.CreateObject("ADODB.Connection")
  connRoom.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacility where numBuildingId = "& numBuildingId &" and strRoomNumber <> '' order by strRoomNumber"
   set rsFillRoom = Server.CreateObject("ADODB.Recordset")
   rsFillRoom.Open strSQL, connRoom, 3, 3
%>
<tr>
  <th>Existing Room Number </th>
  <td>
   <% strRoomNumber = cstr(request.form("cboRoomNumber")) %>
       <select size="1" name="cboRoomNumber" onchange="javascript:FillBuildingCampus()">
         <option value="0"
          <% if strRoomNumber = "" then
    response.Write "select any one"
    end if %>>Select any one</option>
        <%while not rsFillRoom.Eof%>
        <option value="<%=rsFillRoom("strRoomNumber")%>"
        <% if rsFillRoom("strRoomNumber") = strRoomNumber  then
    response.Write "selected"
    end if %>><%=rsFillRoom("strRoomNumber")%></option>
        <%rsFillRoom.Movenext
         wend 
         
         ' closing the connections
         
           rsFillRoom.close
           set rsFillRoom = nothing
           connRoom.Close
           set connRoom = nothing
       %>
       </select></form></td>
 </tr>
 <tr>
  <th><form method="post" action="AdminEdit.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
 <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
    <input type="hidden" name="hdnRoomNumber" value="<%=strRoomNumber%>" />
    <input type="hidden" name="hdnOption" value="Facility" />
	</th>
  <tr>
<%
Dim connRoomName
Dim rsFillRoomName

'Database Connectivity Code 
  set connRoomName = Server.CreateObject("ADODB.Connection")
  connRoomName.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacility where strRoomNumber = '"& strRoomNumber&"'"  
   set rsFillRoomName = Server.CreateObject("ADODB.Recordset")
   rsFillRoomName.Open strSQL, connRoomName, 3, 3
%>

    <% if  rsFillRoomname.EOF <> True then 
       dim strSupName    
       'jan 2010 change to repair relationship facility:supervisor
       'strsupName = rsFillRoomName("strFacilitySupervisor")
       strsupID = rsFillRoomName("numFacilitySupervisorID")
     %>
        <th>New Room Number</th>
        <td><input type="text" name="txtRoomNumber" size="20" value="<%=rsFillRoomName("strRoomNumber")%>" /></td>
       </tr>
      <tr>
        <th>New Room Name</th>
        <td><input type="text" name="txtRoomName" size="20" value="<%=rsFillRoomName("strRoomName") %>" /></td>
      </tr>
     <%else %> 
        <th>New Room Number</th>
        <td><input type="text" name="txtRoomNumber" size="20" value="" /></td>
        
      </tr>
      <tr>
        <th>New Room Name</th>
        <td><input type="text" name="txtRoomName" size="20" value="" /></td>
      </tr>

     <%end if%> 
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
          
          'jan 2010 change to repair relationship facility:supervisor
             'if (strSupName = rsFillSup("strLoginId")) then
             if (strSupID = rsFillSup("numSupervisorID")) then %>
        <!--<option value = "<%=rsFillSup("strLoginId")%>" selected><%= cstr(rsFillsup("strGivenName")) + " " + cstr(rsFillsup("strSurName")) %></Option> -->
        <option value = "<%=rsFillSup("numSupervisorID")%>" selected><%= cstr(rsFillsup("strGivenName")) + " " + cstr(rsFillsup("strSurName")) %></Option> 
       <% else%>
        <!--<option value = "<%=rsFillSup("strLoginId")%>"><%= cstr(rsFillsup("strGivenName")) + " " + cstr(rsFillsup("strSurName")) %></Option> -->
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
		<!-- CL note: This button does not work in Mozilla etc. as it calls VBscript - replaced with a reset button instead 
		<input type="Button" value="Clear" name="btnClear" onclick="clearform()">
		-->
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