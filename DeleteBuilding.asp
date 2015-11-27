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
 <title>Online Risk Register - Delete a Building</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if (document.DeleteBuilding.cboCampusName.value !="0" && document.Form1.cboBuildingName.value !="0") 
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
 document.DeleteBuilding.submit();
}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Delete a Building</h1>
 
 <center>

<form method="post" action="DeleteBuilding.asp" name="DeleteBuilding">
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

<table class="adminfn" style="width: 65%">
<tr>
  <th>Existing Campus Name:</th>
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
  </form>
  <form method="post" action="AdminDelete.asp" name="Form1" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
   <tr>
  <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" /> 
  <input type="hidden" name="hdnOption" value="Building" />
  <th>Existing Building Name:</th>
      <%
' Code to fill the Building names for a particular campus

Dim connBuilding
Dim rsFillBuilding
 ' setting up the database connectivity
 
   set connBuilding = Server.CreateObject("ADODB.Connection")
   connBuilding.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblBuilding where numCampusId ="& numCampusID &" order by strBuildingName"
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, connBuilding, 3, 3
  
  numBuildingID = cint(request.form("cboBuildingName"))
    if numBuildingID = "" then
	   numBuildingID = 0
    end if

%>
      <td>
      <select size="1" name="cboBuildingName">
        <option value="0">Select any one</option>
		  <% while Not rsFillBuilding.Eof %>
		   <option value="<%=rsFillBuilding("numBuildingId")%>"><%=rsFillBuilding("strBuildingName")%></option>
           <%rsFillBuilding.Movenext
            wend 
         
         ' closing the connections
         
           rsFillBuilding.close
           set rsFillBuilding = nothing
           connBuilding.Close
           set connBuilding = nothing
          %>
      </select>
     </td>
    </tr>
<tr>
 <td colspan="2">
 <center>
  <input type="submit" value="Delete" name="btnSave" />
 </center>
 </td>
</tr>
 </table>

</form>

</center>

</div></div>

</body>
</html>