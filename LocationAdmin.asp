<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<%
 if session("strLoginId") <> "admin" then
  response.redirect "AccessRestricted.htm"
 end if
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Online Risk Register - Select a location for creating Risk Assessment</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if((document.Form1.cboBuilding.value == "0" || document.Form2.cboRoom.value == "0") && document.Form2.cboOperation.value == "0") 
  {
     alert ("Please make a selection in all fields on the location form");	
	   return(false);
   }  
  else   
  {
 
      return;
	}
}
// function to reload the form to add the new entries
function ChangeType(val)
{
 document.Form2.QORAtype.value = val;
 //console.log(document.Form2.QORAtype.value);

}
// function to reload the form to add the new entries
function FillDetails()
{
 document.Form1.submit();

}
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content">

 <h1 class="pagetitle">Select a location or operation for creating Risk Assessment</h1>

 <center>
 
 <form method="post" action="LocationAdmin.asp" name="Form1">

<table class="adminfn" style="width: 60%">
<tr><td style="border-bottom: 1px solid rgb(68, 92, 44);"><table>

  <tr>
   <th>Campus:</th>
   <td>


<%'******* code to fill the campus******** %>
<%
Dim connCampus
Dim rsFillCampus
Dim numCampusId

'Database Connectivity Code 
  set connCampus = Server.CreateObject("ADODB.Connection")
  connCampus.open constr
      
   ' setting up the recordset
  strSQL ="Select * from tblCampus"

  set rsFillCampus = Server.CreateObject("ADODB.Recordset")
  rsFillCampus.Open strSQL, connCampus, 3, 3
%>

<%    numCampusID = cint(request.form("cboCampus"))
        if numCampusID = "" then
	       numCampusID = 0
        end if 
   
 %>

 <%'************* Populating Campus*******************************%>
            <select size="1" name="cboCampus" onchange="javascript:FillDetails()">
            <option value= "0"
                <% if numCampusID = 0 then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <%while not rsFillCampus.Eof%>
        <option value="<%=rsFillCampus(0)%>"
        <% if rsFillCampus(0) = numCampusID then
		  response.Write "selected='selected'"
		  end if %>><%=cstr(rsFillCampus(1))%></option>
        <%rsFillCampus.Movenext
         wend 
         
         ' closing the connections
         
           rsFillCampus.close
           set rsCampus = nothing
           connCampus.Close
           set connCampus = nothing
         %>
            </select>
		</td>
	</tr>
	<tr>
		<th>Building:</th>
		<td>
            <%'******* code to fill the Building*********%>
<%
Dim conn
Dim rsFillBuilding

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
        numCampusID = cint(request.form("cboCampus"))
  
   ' setting up the recordset
 
   strSQL ="SELECT Distinct(tblBuilding.numBuildingId),"_
   &" tblBuilding.strBuildingName "_
   &" FROM tblBuilding, tblCampus "_ 
   &" WHERE  "_
   &" tblBuilding.numCampusId=tblCampus.numCampusID and tblCampus.numCampusId ="&numCampusId 
   
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, conn, 3, 3
%>
  <%    numBuildingID = cint(request.form("cboBuilding"))
        if numBuildingID = "" then
	       numBuildingID = 0
        end if 
        
        %>
    <select size="1" name="cboBuilding" onchange="javascript:FillDetails()">
    <option value="0"
         <% if numBuildingID = 0 then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <%while not rsFillBuilding.Eof%>
        <option value="<%=rsFillBuilding(0)%>"
        <% if rsFillBuilding(0) = numBuildingID then
		  response.Write "selected='selected'"
		  end if %>><%=cstr(rsFillBuilding(1))%></option>
        <%rsFillBuilding.Movenext
         wend 
         
         ' closing the connections
         
           rsFillBuilding.close
           set rsFillBuilding = nothing
           conn.Close
           set conn = nothing
         %>
   </select>
   </td>
   </tr>
   </form>
   <tr>
    <th>Room Name/Number:</th>
	<td>
   <form method="post" action="CQORAAdmin.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
	 <input type="hidden" name="hdnBuildingId" value="<%=numBuildingID%>" />
	 <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>"/>
	 <input type="hidden" name="QORAtype" value=""/>
<%'****** Code to fill the Room Name and Room Number *******%>
<%
Dim connR
Dim rsFillR

'Database Connectivity Code 
  set connR = Server.CreateObject("ADODB.Connection")
  connR.open constr
 
 ' setting up the recordset
 
        numCampusID = cint(request.form("cboCampus"))
        numBuildingID = cint(request.form("cboBuilding"))
   ' setting up the recordset
 
   strSQL ="SELECT tblFacility.strRoomNumber,tblFacility.strRoomName,tblFacility.numFacilityId "_
   &" FROM tblFacility, tblBuilding, tblCampus "_ 
   &" WHERE tblFacility.numBuildingID=tblBuilding.numBuildingID "_
   &" And tblBuilding.numCampusId=tblCampus.numCampusID "_
   &" And tblCampus.numCampusId ="&numCampusId &" and tblBuilding.numBuildingId = "& numBuildingId 
   
   set rsFillR = Server.CreateObject("ADODB.Recordset")
   rsFillR.Open strSQL, connR, 3, 3
%>

            <select size="1" name="cboRoom" onchange="javascript:ChangeType('location')">
             <option value="0">Select any one</option><%While not rsFillR.EOF %>
		     <option value="<%=rsFillR(2)%>"><%=cstr(rsFillR(0))+ " - "+cstr(rsFillR(1))%></option>
		     <%
		       rsFillR.Movenext
		       wend
		     %>
      </select>
</td>
</tr>

 </table></td>
<td style="border-left: 1px solid rgb(68, 92, 44);border-bottom: 1px solid rgb(68, 92, 44);">
 <table>
 <tr><th>Operation: </th>
 	<td><select size="1" name="cboOperation" onchange="javascript:ChangeType('operation')">
 	<option value="0">Select any one</option>
 	<% 
 	Dim operConn
 	Dim rsOperations
 	set operConn = Server.CreateObject("ADODB.Connection")
    operConn.open constr
    strSQL = "select * from tblOperations"
    set rsOperations = Server.CreateObject("ADODB.Recordset")
    rsOperations.Open strSQL, operConn, 3, 3
    While not rsOperations.EOF %>
		<option value="<%=rsOperations("numOperationId")%>"><%=cstr(rsOperations("strOperationName"))%></option>
		<%
		rsOperations.Movenext
	wend
 	%>
 	</select>
 	</td></tr>
 </table>
 </td></tr>


<tr>
<td colspan="2"><center>
<input type="submit" value="Next" name="btnProceed" />
</form>
</center>
  </td>
 </tr>
 

 
 

 </table>
 

</center>

</div></div>

</body>
</html>