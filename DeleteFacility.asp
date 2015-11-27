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
 
   strSQL ="Select distinct(tblFacility.numBuildingId)as NumBuildingID,tblCampus.strCampusName,tblBuilding.strBuildingName from tblBuilding,tblCampus,tblFacility where tblBuilding.numCampusId = tblCampus.numCampusId and tblFacility.numBuildingId = tblBuilding.numBuildingId order by strBuildingName"
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
function FillBuildingCampus()
{
 document.DeleteFacility.submit();

}
// to reset the form
function resetForm()
{
	document.deletefacility.cboBuildingName.selectedIndex = 0
}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Delete a Facility</h1>
 
 <center>

<table class="adminfn" style="width: 55%">
<form method="post" action="DeleteFacility.asp" name="DeleteFacility">
<tr>
 <th>Existing Building Name</th>
 <td>

  <%    numBuildingID = cint(request.form("cboBuildingName"))
        if numBuildingID = "" then
	       numBuildingID = 0
        end if %>
        <select size="1" name="cboBuildingName" onChange="javascript:FillBuildingCampus()">
        <option value="0"
         <% if numBuildingID = 0 then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <%while not rsFillBuilding.Eof%>
        <option value="<%=rsFillBuilding("NumBuildingID")%>"
        <% if rsFillBuilding("NumBuildingID") = numBuildingID then
		  response.Write "selected='selected'"
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
      <tr>
<%
Dim connRoom
Dim rsFillRoom

'Database Connectivity Code 
  set connRoom = Server.CreateObject("ADODB.Connection")
  connRoom.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacility where numBuildingId = "& numBuildingId &" and strRoomNumber <> '' order by strRoomNumber "
   set rsFillRoom = Server.CreateObject("ADODB.Recordset")
   rsFillRoom.Open strSQL, connRoom, 3, 3
%>
        <th>Existing Room Number:</th>
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
		  response.Write "selected='selected'"
		  end if %>><%=rsFillRoom("strRoomNumber")%></option>
        <%rsFillRoom.Movenext
         wend 
         
         ' closing the connections
         
           rsFillRoom.close
           set rsFillRoom = nothing
           connRoom.Close
           set connRoom = nothing
         %>
       </select></td>
      </tr>
      
    </form>
    
    <form method="post" action="AdminDelete.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice(this);">
    <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />  
    <input type="hidden" name="hdnRoomNumber" value="<%=strRoomNumber%>" />  
    <input type="hidden" name="hdnOption" value="Facility" />  

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
       'AA jan 2010 repair relationship alter
       'strsupName = rsFillRoomName("strFacilitySupervisor")
       strsupName = rsFillRoomName("numFacilitySupervisorID")

     %>
      
      <tr>
        <th>Existing Room Name:</th>
        <td><input type="text" name="txtRoomName" size="40" value="<%=rsFillRoomName("strRoomName") %>" /></td>
      </tr>
      
     <%else %> 
     
      <tr>
        <th>Existing Room Name:</th>
        <td><input type="text" name="txtRoomName" size="40" value="" /></td>
      </tr>

     <%end if%> 
<%
Dim connSup
Dim rsFillSup

'Database Connectivity Code 
  set connSup = Server.CreateObject("ADODB.Connection")
  connSup.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySuperVisor" 
   set rsFillSup = Server.CreateObject("ADODB.Recordset")
   rsFillSup.Open strSQL, connSup, 3, 3
%>
  
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