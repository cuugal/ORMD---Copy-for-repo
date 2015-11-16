<%@Language = VBscript%>
<% session("My_Session") = "Open" %>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">

<%
'Set this session variable to prevent displaying of dates in american format (Access has no proviso for normal dates, only US)
session.LCID = 2057	'English(British) format

Dim connFac
Dim rsFillFac
Dim strSQL
' VERSION number on or about line 86
'Database Connectivity Code 

  set connFac = Server.CreateObject("ADODB.Connection")
  connFac.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFaculty order by strFacultyName"
   set rsFillFac = Server.CreateObject("ADODB.Recordset")
   rsFillFac.Open strSQL, connFac, 3, 3
%>
<!-- <head> CL removed extraneous markup 1/11/2010-->
<script type="text/javascript">

// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
       answer = confirm("Are you sure you want to save these changes?")
		if (answer == true) 
		{ 
           return ;
		} 
		else
		 { 
		 return (false);
		}  
}
// function to reload the form to add the new entries
function FillDetailsSuper()
{
 	document.MenuSuper.submit();
}
function FillDetailsLocation()
{
 	document.MenuLocation.submit();
}
function FillDetailsOperation()
{
 	document.MenuOperation.submit();
}
function FillDetailsTask()
{
 	document.MenuTask.submit();
}

function FillSearch(){
	return;
}

//Function to clear the contents of the form
function resetForm()
{ 
  document.Menu.txtHazardousTask.Value = "*"
}

function clearform()
{
/*
  document.MenuSuper.reset();
  document.MenuOperation.reset();
  document.MenuLocation.reset();
  document.MenuTask.reset();
  */
  var str 
  str = "Menu.asp";
  window.location.replace(str); 
}


//This focuses the user client on the login ID form, for ease of login - CL 9/5/2011
function formfocus() {
   document.getElementById('txtLoginID').focus();
   }
 window.onload = formfocus;
</script>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!-- Generic Metadata -->
 <meta name="description" content="The Online Risk Register of the University of Technology, Sydney." />
 <meta name="keywords" content="online,risk,register,environment,health,safety,assessment,database,branch,ohs,occupational,university,technology,sydney,australia, hazards,first,aid,oms,online,management,system,responsibilities,legal" />
 <meta name="author" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
 <!-- Dublin Core Metadata -->
 <meta name="dc.title" content="Online Risk Register - Safety &amp; Wellbeing, University of Technology, Sydney" />
 <meta name="dc.subject" content="risk,register,assessment,health,safety,whs,occupational,university,technology,sydney,hazards" />
 <meta name="dc.description" content="The Online Risk Register of the University of Technology, Sydney." />
 <meta name="dc.identifier" scheme="URL" content="http://www.safetyandwellbeing.uts.edu.au" />
 <meta name="dc.creator" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
 <meta name="dc.creator.Address" content="safetyandwellbeing@uts.edu.au" />
 <meta name="dc.language" scheme="RFC1766" content="en" />
 <meta name="dc.publisher" content="Safety &amp; Wellbeing, University of Technology, Sydney" />
 <title>Online Risk Register - Search health and safety risk assessments</title>
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
<base target="_self" />
</head>

<body>

<div id="wrapper">

<div id="orrheading"><h1>Online Risk Register for Facilities and Operations / Projects</h1><img src="http://www.uts.edu.au/images/css/utslogo.gif" alt="UTS" width="132" height="30" style="border: none; vertical-align: middle; float: right; " /><font size="1" color="gray">V2.0</font></div>
<!-- Version 1.3 30 June 2011 -->

<%'*********************Login component********************************************************%>

<%
Response.Expires=0
Response.Buffer = True

If Request.Form ("btnLogin") = "Login" then

Dim strLoginID, strPassword, msg
dim strAccessLevel
Dim strSQL2
Dim rsAccess2
Dim conn2 

strLoginID=request.form("txtLoginID")
strPassword=request.form("txtPassword")
strSQL2="select * from tblFacilitySupervisor where strLoginID='"& strLoginID &"'"



set conn2 = Server.CreateObject("ADODB.Connection")
conn2.open constr
set rsAccess2 = Server.CreateObject("ADODB.Recordset")
rsAccess2.Open strSQL2, conn2, 3, 3
If  rsAccess2.eof then
   'msg = "The login ID: " + strLoginID + " does not exist, try a different Login ID. Please contact administrator if you need a login ID and Password"
   'Response.Write(msg) %> 
 <script type="text/javascript">
   alert("Invalid username or password , Please Try again !")
 </script>
   
<%else If  rsAccess2("strPassword")= strPassword then
		session("LoggedIn")= true
		session("strLoginID")= strLoginID
		strAccessLevel = rsAccess2("strAccessLevel")
		session("strAccessLevel") = strAccessLevel
		
		'get the username & put into session data to avoid annoying timeout message
		set conn3 = Server.CreateObject("ADODB.Connection")
  		conn3.open constr
  		strSQL = "Select strGivenName,strSurname "_
  		&" from tblfacilitySupervisor"_
  		&" where tblFacilitySupervisor.strLoginId = '"& strLoginID &"'" 
  
  		set rsSearchLogin = server.CreateObject("ADODB.Recordset")
  		rsSearchLogin.Open strSQL, Conn3, 3, 3
  		
  		strName = cstr(rsSearchLogin(0)) + " " + cstr(rsSearchLogin(1))
  		session("strName") = strName
		
		if rsAccess2("strAccessLevel") ="A" then
		   Response.Redirect "indexLoggedAdmin.asp"
		elseif rsAccess2("strAccessLevel") ="S" then
		   Response.Redirect "indexLoggedSupervisor.htm"
		   elseif rsAccess2("strAccessLevel") ="D" then
		   Response.Redirect "indexLoggedDD.htm"
		end if   
		
	else
		'msg = "The password for " + strLoginID + " was not correct, please try again. Please contact administrator if you need a new Password"
		'Response.Write(msg)%>
		 <script type="text/javascript">
           alert("Invalid username or password - please try again.")
         </script>
   <%
	end if
end if

else
	msg = Request.QueryString("msg")
	if msg = "noaccess" then
		'msg = "Please login, your user session has timed out (or you have not logged in yet)."
		'Response.Write(msg)
		%>
		 <script language = javascript>
			alert("Your session has timed out - please re-login.")
		</script>
   
		<%
	end if
End if

%>
<%'***********************end login component **********************************************************%>

<div id="content">

	<h1 class="pagetitle">Search Risk Assessments</h1>
	
	<table>
	<tr><td valign="top">
	<div style="align:top; float:left;display:inline;padding-left:40px;padding-right:40px">
		<form action="Menu.asp" name="frmlogin" method="post">
		<table class="bluebox">
		<h2>Login</h2>
		<tr><td>Username:</td></tr>
		<tr>
   			<td><input name="txtLoginID" id="txtLoginID" maxlength="70" type="text" value="<%=strLoginID%>" size="25" /></td>
		</tr>
		<tr><td>Password:</td></tr>
		<tr>
   			<td><input name="txtPassword" maxlength="70" type="password" size="25" /></td>
		</tr>
		<tr>
			<td align="right"><input type="submit" value="Login" name="btnLogin" />&nbsp;&nbsp;&nbsp;<input type="reset" value="Clear" name="btnReset" /></td>
		</tr>
		<tr>
			<td><h5>Note:<br/> Username and password <br/> are case sensitive.</h5></td>
		</tr>
		<tr>
			<td><h5>To obtain login details <br /> contact Safety and <br /> Wellbeing on ext 1063.</h5></td>
		</tr>
		</table>
		</form>
	</div>
	</td>
	<td>


	<div style="float:left; display:inline;">
<!-- old code is below -->

	<table class="searchtable" style="padding-top:15px">
<%'********************************** SEARCH SUPERVISOR  **************************************************************%>		
		<form method="post" action="Menu.asp" name="MenuSuper">
			<tr>
				<td colspan="4">Search by <B>Supervisor</B> OR by <B>Facility Location</B> OR by <B>Operation</B> OR by <B>Risk Assessment Number</B></td>
			</tr>
			<tr>
				<td colspan="4"><hr /></td>
			</tr>
			<tr>
				<td>Search <b>Supervisors</b></td>
			</tr>
			<tr>
				<th>Faculty/Unit</th>
					<td>
	   <%numFacultyID = cint(request.form("cboFacultySuper"))
			if numFacultyID = "" then
			   numFacultyID = 0
			end if %>
				
				<select size="1" name="cboFacultySuper" tabindex="1" onChange="javascript:FillDetailsSuper()">
				<option value="0"
			<% if numFacultyID = 0 then
			response.Write "Select any one"
			end if %>
			>Select any one</option>
				<%while not rsFillFac.Eof 
					if rsFillFac("boolActive")= True Then %>
						<option value="<%=rsFillFac("NumFacultyID")%>"
							<% if rsFillFac("NumFacultyID") = numFacultyID Then
							response.Write "selected"
							end if %>
						><%=cstr(rsFillFac("strFacultyName"))%></option>
					<% End If
					rsFillFac.Movenext	
				wend 
			 
			 %>
			</select></td>
			</tr>
	</form>
	<form method="post" name="Submit1" action="CollectInfo.asp" name = "f1" enctype = "application/x-www-form-urlencoded" >
	  <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
	  <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
	  <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
	  <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
	  <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
	  <input type="hidden" name="cboFaculty" value="<%=cboFacultySuper%>" />
	  <input type="hidden" name="searchType" value="supervisor" />
			<tr>
				<th>Supervisor Name</th>
				<td>
	<%'******* code to fill the Supervisor*****%> 
	<%
	Dim connSup
	Dim rsFillSup
	Dim strSuperv

	'Database Connectivity Code 
	  set connSup = Server.CreateObject("ADODB.Connection")
	  connSup.open constr
	 
	 ' setting up the recordset
	 
	   strSQL ="Select * from tblFacilitySupervisor where numFacultyId ="&numFacultyId &" order by strGivenName "
	   set rsFillSup = Server.CreateObject("ADODB.Recordset")
	   rsFillSup.Open strSQL, connSup, 3, 3
	%>
	<%
		   strSuperv = request.form("cboSupervisorName")
		   
	%>
		  <select size="1" name="cboSupervisorName" tabindex="2" >
			<option value="0" 
			  <% if strSuperV = "" then
				 response.Write "select any one"
			  end if %>>Select any one</option>


			  <%while not rsFillSup.Eof
                  if rsFillSup("boolDeprecated") = 0 then%>
					<option value="<%=rsFillSup("strLoginID")%>"
						<% if rsFillSup("strLoginId") = strSuperV   then
							response.Write "selected"
							end if %>
						><%=cstr(rsFillSup("strGivenName")) + " " + cstr(rsFillSup("strsurname")) %></option>
					<% 
					End if
					rsFillSup.Movenext
			   wend 

			 ' closing the connections
			   rsFillSup.close
			   set rsFillSup = nothing
			   connSup.Close
			   set connSup = nothing
			 %>

				</select>&nbsp;</td>
				</tr>
			
			
			  	<tr>
	   				<td colspan="2"><center><input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /--> <!--DLJ Removed this button from common search 22July2011-->
	   		</form></center></td>
	  </tr>

			

<%'************************************************ END SEARCH SUPERVISOR ******************************************************** %>
	<tr>
	   <td colspan="3"><hr /></td><td>OR</td>
	</tr>
	
<%'************************************************  SEARCH LOCATION ******************************************************** %>	  
	  
	  <form method="post" action="Menu.asp" name="MenuLocation">
			<tr><td>Search <b>Facility Locations</b></td></tr>
				<tr><th>Faculty/Unit</th>
				<td>
	   <%    numFacultyID = cint(request.form("cboFacultyLocation"))
			if numFacultyID = "" then
			   numFacultyID = 0
			end if %>
				
				<select size="1" name="cboFacultyLocation" tabindex="1" onChange="javascript:FillDetailsLocation()">
			<option value="0"
			<% if numFacultyID = 0 then
			response.Write "Select any one"
			end if %>
			>Select any one</option>
				<%rsFillFac.MoveFirst
				while not rsFillFac.Eof 
						'DLJ put this if statement in 22 Jan 2010 - is this OK?
						if rsFillFac("boolActive")= True Then %>
							<option value="<%=rsFillFac("NumFacultyID")%>"
								<% if rsFillFac("NumFacultyID") = numFacultyID Then
								response.Write "selected"
								end if %>
							><%=cstr(rsFillFac("strFacultyName"))%></option>
						<% End If
						rsFillFac.Movenext	
				 wend 
			 
			 %>
			</select></td>
			</tr>
			
	   <tr>
	   
	   <th >Building</th>
	   <%'******* code to fill the Building*****%>
	<%
	Dim conn
	Dim rsFillBuilding

	'Database Connectivity Code 
	  set conn = Server.CreateObject("ADODB.Connection")
	  conn.open constr
	 
	 ' setting up the recordset
	 
		   strSuperv = request.form("cboSupervisorName")
			numCampusID = cint(request.form("cboCampus"))
			'response.write(numCampusId)
		   
	 
	   strSQL = "Select distinct(tblFacility.numBuildingId)as NumBuildingID,tblCampus.strCampusName,tblBuilding.strBuildingName "_
	   &"from tblBuilding,tblCampus,tblFacility, tblFacilitySupervisor, tblFaculty "_
	   &"where tblFaculty.numFacultyID="& numFacultyID&" "_
	   &"and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultYID "_
	   &"and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID "_
	   &"and tblFacility.numBuildingId = tblBuilding.numBuildingId "_
	   &"and tblBuilding.numCampusId = tblCampus.numCampusId "_
	   &" order by strBuildingName"
	   
	   'response.write(strSQL)
	 'end if
	   
	   set rsFillBuilding = Server.CreateObject("ADODB.Recordset")
	   rsFillBuilding.Open strSQL, conn, 3, 3
	%>
	  <%    numBuildingID = cint(request.form("cboBuilding"))
			if numBuildingID = "" then
			   numBuildingID = 0
			end if 
			
			%>
			<td>
			<select size="1" name="cboBuilding" tabindex="4" onChange="javascript:FillDetailsLocation()">
			 <option value="0"
			 <% if numBuildingID = 0 then
			  response.Write "select any one"
			  end if %>>Select any one</option>
			<%while not rsFillBuilding.Eof%>
			<option value="<%=rsFillBuilding("numBuildingID")%>"
			<% if rsFillBuilding("numBuildingID") = numBuildingID then
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
		</form>	
		<form method="post" name="Submit2" action="CollectInfo.asp" name = "f1" enctype = "application/x-www-form-urlencoded" >
			  <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
			  <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
			  <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
			  <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
			  <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
			  <input type="hidden" name="cboFaculty" value="<%=cboFacultyLocation%>" />
			  <input type="hidden" name="searchType" value="location" />
	 <tr>
	  <th>Room No. / Name</th>
	   <%'******Code to fill the Room Name and Room Number****%>
	<%
	Dim connR
	Dim rsFillR

	'Database Connectivity Code 
	  set connR = Server.CreateObject("ADODB.Connection")
	  connR.open constr
	 
	 ' setting up the recordset
	 numCampusID = cint(request.form("cboCampus"))
	 numBuildingID = cint(request.form("cboBuilding"))

	   strSQL ="SELECT tblFacility.strRoomNumber,tblFacility.strRoomName,"_
	   &" tblBuilding.strBuildingName,tblFacility.numFacilityId, strGivenName, strSurname"_
	   &" FROM tblFacility, tblBuilding, tblFacilitySupervisor , tblFaculty"_ 
	   &" WHERE tblFacility.numBuildingID=tblBuilding.numBuildingID "_
	   &" and tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
	   &" and tblFaculty.numFacultyID = tblFacilitySupervisor.numFacultyID "_
	   &" and tblFaculty.numFacultyID = "& numFacultyID&" "_
	   &" And  tblBuilding.numBuildingId = "& numBuildingId &" "_
	   &" order by tblFacility.strRoomNumber"

 
	   set rsFillR = Server.CreateObject("ADODB.Recordset")
	   rsFillR.Open strSQL, connR, 3, 3
	%>
				<td><select size="1" name="cboRoom" tabindex="5" >
				 <option value="0">Select any one</option>
				 <%While not rsFillR.EOF 
				 if len(strSuperv) <= 1 then
				 	facility_name =cstr(rsFillR("strRoomNumber"))+ " - "+cstr(rsFillR("strRoomName"))&" - "&rsFillR("strGivenName")&" "&rsFillR("strSurname")
				 else
				 	facility_name =cstr(rsFillR("strRoomNumber"))+ " - "+cstr(rsFillR("strRoomName"))
				 end if	
				 	%>
				 <option value="<%=rsFillR("numFacilityId")%>"><%=facility_name%></option>
				 <%
				   rsFillR.Movenext
				   wend
				 %>
				 </select></td>
				</tr>
		<tr>
	   <td colspan="2"><center><input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;
	   </form></center></td>
	  </tr>
<%'************************************************ END SEARCH LOCATION ******************************************************** %>
		<tr>
	   	<td colspan="3"><hr /></td><td>OR</td>
	  	</tr>
<%'************************************************  SEARCH OPERATION ******************************************************** %>
	

	<form method="post" action="Menu.asp" name="MenuOperation">
				<tr><td>Search <b>Operations/Projects</b></td></tr>
				<tr><th>Faculty/Unit</th>
				<td>
	   <%    numFacultyID = cint(request.form("cboFacultyOperation"))
			if numFacultyID = "" then
			   numFacultyID = 0
			end if %>
				
				<select size="1" name="cboFacultyOperation" tabindex="1" onChange="javascript:FillDetailsOperation()">
			<option value="0"
			<% if numFacultyID = 0 then
			response.Write "Select any one"
			end if %>
			>Select any one</option>
				<%rsFillFac.MoveFirst
				while not rsFillFac.Eof 
						'DLJ put this if statement in 22 Jan 2010 - is this OK?
						if rsFillFac("boolActive")= True Then %>
							<option value="<%=rsFillFac("NumFacultyID")%>"
								<% if rsFillFac("NumFacultyID") = numFacultyID Then
								response.Write "selected"
								end if %>
							><%=cstr(rsFillFac("strFacultyName"))%></option>
						<% End If
						rsFillFac.Movenext	
				 wend 
			 
			 %>
			</select></td>
			</tr>
	<%	   
	   strSQL = "Select numOperationID, strOperationName , strGivenName, strSurname from tblOperations, tblFacilitySupervisor, tblFaculty where tblFaculty.numFacultyID="& numFacultyID&" and tblFacilitySupervisor.numSupervisorId = tblOperations.numFacilitySupervisorId and tblFaculty.numFacultyId = tblFacilitySupervisor.numFacultyID order by strOperationName"
	 'end if
	   
	   set rsFillOperation = Server.CreateObject("ADODB.Recordset")
	   rsFillOperation.Open strSQL, connR, 3, 3
	%>
	</form>
<form method="post" name="Submit3" action="CollectInfo.asp" name = "f1" enctype = "application/x-www-form-urlencoded" >
	  <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
	  <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
	  <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
	  <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
	  <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
	  <input type="hidden" name="cboFaculty" value="<%=cboFacultyOperation%>" />
	  <input type="hidden" name="searchType" value="operation" />
	

	<tr>
		<th>Operation</th>
			<td><select name="cboOperation" >
				<option value="0">Select any one</option>
				<%while not rsFillOperation.Eof
				if len(strSuperv) <= 1 then
				 	operation_name =rsFillOperation("strOperationName")&" - "&rsFillOperation("strGivenName")&" "&rsFillOperation("strSurname") 
				 else
				 	operation_name =rsFillOperation("strOperationName")
				 end if	
				%>
				<option value="<%=rsFillOperation("numOperationID")%>"
					><%=operation_name%></option>
					<%rsFillOperation.Movenext
			 	wend
			 	rsFillOperation.close
			   set rsFillOperation = nothing
			   connR.Close
			   set connR = nothing
			 	%>
				</select></td>
		</tr>
		<tr>
	   <td colspan="2"><center><input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /--> <!--DLJ Removed this button from common search 22July2011-->
	   </form></center></td>
	  </tr>
<%'************************************************ END SEARCH OPERATION ******************************************************** %>
		
	  <tr>
	   <td colspan="3"><hr /></td><td>OR</td>
	  </tr>
<%'************************************************  SEARCH TASK ******************************************************** %>
	<form method="post" action="Menu.asp" name="MenuTask">
				<tr><td>Search <b>Risk Assessment Number</b></td></tr>
				<tr><th>Faculty/Unit</th>
				<td>
	   <%    numFacultyID = cint(request.form("cboFacultyTask"))
			if numFacultyID = "" then
			   numFacultyID = 0
			end if %>
				
				<select size="1" name="cboFacultyTask" tabindex="1" onChange="javascript:FillDetailsTask()">
			<option value="0"
			<% if numFacultyID = 0 then
			response.Write "Select any one"
			end if %>
			>Select any one</option>
				<%rsFillFac.MoveFirst
				while not rsFillFac.Eof 
						'DLJ put this if statement in 22 Jan 2010 - is this OK?
						if rsFillFac("boolActive")= True Then %>
							<option value="<%=rsFillFac("NumFacultyID")%>"
								<% if rsFillFac("NumFacultyID") = numFacultyID Then
								response.Write "selected"
								end if %>
							><%=cstr(rsFillFac("strFacultyName"))%></option>
						<% End If
						rsFillFac.Movenext	
				 wend 
			 
			 %>
			</select></td>
			</tr>
			</form>
	<form method="post" name="Submit4" action="CollectInfo.asp" name = "f1" enctype = "application/x-www-form-urlencoded" >
	  <input type="hidden" name="hdnSuperV" value="<%=strsuperV%>" />
	  <input type="hidden" name="hdnHazardousTask" value="<%=strHazardousTask%>" />
	  <input type="hidden" name="hdnBuildingId" value="<%=numBuildingId%>" />
	  <input type="hidden" name="hdnCampusID" value="<%=numCampusId%>" />
	  <input type="hidden" name="hdnFacultyId" value="<%=numFacultyId%>" />
	  <input type="hidden" name="cboFaculty" value="<%=cboFacultyTask%>" />
	  <input type="hidden" name="searchType" value="task" />	
	  
	
	   <tr><th>Task/RA Number</th>
	   <td><input type="text" name="txtHazardousTask" size="40" tabindex="0" /></td>
	</tr>
	<tr><td></td></tr>
	  <tr>
	   <td colspan="2"><center><input type="Submit" value="Search" name="btnSearch" onclick="FillSearch()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="Clear Form" name="btnClear" onclick="clearform()" />&nbsp;&nbsp;&nbsp;&nbsp;<!--input type="Submit" value="Action Status Report" name="btnSearch" onclick="FillSearch()" /--> <!--DLJ Removed this button from common search 22July2011-->
	   </form></center></td>
	  </tr>
	 </table>
<%'************************************************ END SEARCH TASK ******************************************************** %>
</center>

<br />

<center>
	<div class="loginlist">
	<ul><li>&nbsp</li>
	</ul>
	<ul>
		<li><a target="_blank" href="UseORRInstruct.htm" title="Read the Online Risk Register documentation.">Instructions for Using this Risk Register</a></li>
		<li><a target="_blank" href="ORRAssessmentInstructions.pdf" title="Read Risk Assessment Instructions for Printing. ">Instructions for Documenting a Risk Assessment (PDF)</a></li>
	</ul>
	<ul<li>&nbsp;</li></ul>
	</div>
<%'********************%>

</div>
<div style="clear:both"></div> 
</td></tr>
</table>
 </div><!-- close the content DIV -->

</div><!-- close the wrapper div -->
</body>
</html>