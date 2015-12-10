<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If

%>
<%
dim strSuperV
strSuperV = Session("strLoginId")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta http-equiv="Content-Language" content="en-au" />
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
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
function FillDetails()
{
 document.Form1.submit();

}
function ChangeType(val)
{
 document.Form2.QORAtype.value = val;
 //console.log(document.Form2.QORAtype.value);

}
</script>
<title>Online Risk Register - Select a location or an operation for creating the Risk Assessment</title>

     <!--#include file="bootstrap.inc"--> 
</head>
<body>
    <!--#include file="HeaderMenu.asp" -->

<% 'empty out previous session:
session("LastRACreatednumFacilityID") = "" 
	session("LastRACreatednumBuildingID") = ""
    session("LastRACreatednumCampusID") = ""
    %>
<div id="wrapper">
  <div id="content">
    <h1 class="pagetitle">Select a location or operation for creating the Risk Assessment</h1>
    <center>
      <form method="post" action="LocationSup.asp" name="Form1">
        <table class="adminfn" style="width: 50%">
        <tr><td style="border-bottom: 1px solid rgb(68, 92, 44);"><table>
        
        <tr>
          <th>Building:</th>
          <td><%'*********************** code to fill the Building**************%>
            <%
Dim conn
Dim rsFillBuilding

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
        numCampusID = cint(request.form("cboCampus"))
  
   ' setting up the recordset
 'AA jan 2010 reworked this to support the relationshsip changes (join to tblFacilitySupervisor)
   strSQL ="SELECT Distinct(tblBuilding.numBuildingId),"_
   &" tblBuilding.strBuildingName "_
   &" FROM tblBuilding, tblfacility ,tblFacilitySupervisor"_ 
   &" WHERE tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and"_
   &" tblBuilding.numBuildingId=tblfacility.numBuildingID and tblFacilitySupervisor.strLoginID ='"& strSuperV &"'" 
   
   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
   rsFillBuilding.Open strSQL, conn, 3, 3

%>
            <%    numBuildingID = cint(request.form("cboBuilding"))
        if numBuildingID = "" then
	       numBuildingID = 0
        end if 

        %>
            <select size="1" name="cboBuilding" id="cboBuilding" onchange="javascript:FillDetails()">
              <option value="0" 
         <% if numBuildingID = 0 then
		  response.Write "select any one"
		  end if %>>Select any one</option>

				<%while not rsFillBuilding.Eof%>

				<option value="<%=rsFillBuilding(0)%>"
					<% If (cstr(rsFillBuilding(0)) = session("LastRACreatednumBuildingID")) or rsFillBuilding(0) = numBuildingID then
					response.Write " selected"
					end if %>>
					<%=cstr(rsFillBuilding(1))%></option>
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
      <form method="post" action="cqoraSup.asp" enctype="application/x-www-form-urlencoded" name="Form2" onsubmit="return ConfirmChoice();">
        <input type="hidden" name="hdnBuildingId" value="<%=numBuildingID%>" />
        <input type="hidden" name="hdnCampusId" value="<%=numCampusID%>" />
        <input type="hidden" name="QORAtype" value=""/>
        <%'*****Code to fill the Room Name and Room Number****%>
        <%
	
Dim connR
Dim rsFillR

'Database Connectivity Code 
  set connR = Server.CreateObject("ADODB.Connection")
  connR.open constr
 
 ' setting up the recordset
 
        numCampusID = cint(request.form("cboCampus"))
        numBuildingID = cint(request.form("cboBuilding"))

		If session("LastRACreatednumBuildingID") <> "" then
			numBuildingID = session("LastRACreatednumBuildingID")
			' Reset session variable to allow user to make different selection if they choose
			session("LastRACreatednumBuildingID") = ""					 
		end if 
		
		
   ' setting up the recordset
 'AA jan 2010 reltn fix: alrered this to include join to tblFacilitySupervisor
   strSQL ="SELECT tblFacility.strRoomNumber,tblFacility.strRoomName,tblFacility.numFacilityId "_
   &" FROM tblFacility, tblBuilding , tblFacilitySupervisor"_ 
   &" WHERE tblFacility.numBuildingID=tblBuilding.numBuildingID "_
   &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID "_
   &"  and tblBuilding.numBuildingId = "& numBuildingId &" and tblFacilitySupervisor.strLoginID = '"& strSuperV&"'" 
   
   set rsFillR = Server.CreateObject("ADODB.Recordset")
   rsFillR.Open strSQL, connR, 3, 3
%>
        <th>Room Name/Number:</th>
        <td><select size="1" name="cboRoom" id="cboRoom" onchange="javascript:ChangeType('location')">
            <option value="0">Select any one</option>
            <%While not rsFillR.EOF %>
            <option value="<%=rsFillR(2)%>"
			<% if cstr(rsFillR(2)) = session("LastRACreatednumFacilityID") then
		  response.Write " selected"
		  '		 Reset session variable to allow user to make different selection if they choose
		  		session("LastRACreatednumFacilityID") = ""
		  end if %>
		  ><%=cstr(rsFillR(0))+ " - "+cstr(rsFillR(1))%>
			</option>
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
    strSQL = "select * from tblOperations, tblFacilitySupervisor where tblFacilitySupervisor.strLoginId = '"&strSuperV&"' and tblOperations.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID"
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
            </center>
      </form>
      </td>
      </tr>
      </table>
    </center>
  </div>
</div>
</body>
</html>
