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

<%dim loginId
loginId = session("strLoginId")%>


<%dim hTaskDesc
  dim hNavigCnt
  
  hTaskDesc = Request.Form("hdnTaskDesc") 
  hNavigCnt = Request.Form("hdnNavigationCnt")
  if hNavigCnt = "1" then
  %>
  <b>The risk assessment named &quot;<%=hTaskDesc%>&quot; has been added successfully.
  <%
  end if
  
  dim pn 
  dim htd
  
  pn = session("pn")
  htd = session("HTask")
  
  if pn = "1" then
    %>
  <span style="font-size: 9pt; color: #000;">The risk assessment named &quot;<%=htd%>&quot; has been edited successfully.</span>
  <%
  end if
  
%>

<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - My Risk Assessments - Administrator</title>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ((document.Form1.cboFacultyName.value == "0") ) 
  {
      alert ("Please select a Faculty/Unit to proceed.");
	   return(false);
	}
  else if ((document.Form1.cboSupervisor.value != "0") && (document.Form2.cboFacility.value =="0"))
  {
      alert ("You need to select both a supervisor and their respective facility.");
	   return(false);
  }	
  
}
// function to reload the form to add the new entries
function FillBuildingCampus()
{
 document.EditFacility.submit();
}

// function to reload the form to add the new entries
function FillFaculty()
{
 document.Form1.submit();
}
</script>

<%
 dim rsSearchFacility1
 dim rsSearchFacility2
 dim rsSearchFaculty
 dim rsSearchFacultyDD
 dim Conn
 dim strSQL
 
   '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr %>
        

 <%  '------------------------get the faculty for the login ---------------
 ' strSQL = "Select * "_
 ' &" from tblfacilitySupervisor,tblFaculty "_
 ' &" where tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyId "_
 ' &" and tblFacilitySupervisor.strLoginId = '"& loginId &"'" 
  
 ' set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  'Response.Write(strSQL) 
 ' rsSearchFaculty.Open strSQL, Conn, 3, 3     
  'strFacultyName = rsSearchFaculty(7)     
 ' strGivenName = rsSearchFaculty(3)
  'strSurname = rsSearchFaculty(4)
'  strName = cstr(strGivenName) + " " + cstr(strSurname)
  %>
</head>

<body>

<center>

<table class="myqora" style="width: 60%;">
<thead>
<tr>
 <th colspan="2">My RAs</th>
</tr>
</thead>
<tbody>
<tr>
 <form method="post" action="MyQoraAdmin.asp" name="Form1">
 <%strSQL = "Select * from tblfaculty"
  
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty.Open strSQL, Conn, 3, 3 %>
  <%    numFacultyID = cint(request.form("cboFacultyName"))
        if numFacultyID = "" then
       numFacultyID = 0
        end if %>
  <th>Faculty/Unit</th>
  <td>
   <select size="1" name="cboFacultyName" onchange="javascript:FillFaculty()">
   <option value="0"
    <% if numFacultyID = 0 then
				 response.Write "select any one"
			end if %>>Select any one</option>
             <%while not rsSearchFaculty.Eof%>
             <option value="<%=rsSearchFaculty("NumFacultyID")%>"
        <% if rsSearchFaculty("NumFacultyID") = numFacultyID then
		  response.Write "selected='selected'"
		  end if %>><%=cstr(rsSearchFaculty("strFacultyName")) %></option>
        <%rsSearchFaculty.Movenext
         wend 
         
         ' closing the connections
         
           rsSearchFaculty.close
           set rsSearchFaculty= nothing
           'conn.Close
           'set conn = nothing
         %>
  <%strSQL = "Select * from tblfaculty where numFacultyID ="& numFacultyID
  
  set rsSearchFacultyDD = server.CreateObject("ADODB.Recordset")
  rsSearchFacultyDD.Open strSQL, Conn, 3, 3 %>
  </select>
  </td>
  </tr>
<!-- Dean/Director extraction code not working - noted by CL 26-6-2006
<tr>
   <th>Dean/Director:</th>
   <td>
   <%
	if rsSearchFacultyDD.EOF <> True then
	 response.write(rsSearchFacultyDD("strDGivenName")+" "+rsSearchFacultyDD("strDsurName"))
	end if %>
   </td>
   </tr> -->
  <tr>
   <th>Supervisor:</th>
   <td>
	<%
	   strLoginID = Request.Form("cboSupervisor") 
	   if strLoginID = "" then
	   strLoginID = ""
	   end if

  strSQL = "Select * from tblfacilitySupervisor where numFacultyID = "& numFacultyID
  set rsSearchSup = server.CreateObject("ADODB.Recordset")
  rsSearchSup.Open strSQL, Conn, 3, 3 %>
  
	<select size="1" name="cboSupervisor" onchange="javascript:FillFaculty()">
	<option value="0" 
			<% if strLoginId = "" then
					response.Write "select any one"
			 end if %>>Select any one</option>
			 
			<%while not rsSearchsup.EOF    %>
			<option value= "<%=rsSearchsup(0)%>" 
			<% if rsSearchsup("strloginID") = strLoginID then
		          response.Write "selected='selected'"
		       end if %>><%=cstr(rsSearchsup(3))+"  "+cstr(rsSearchsup(4)) %></option>
			<%rsSearchsup.MoveNext 
			wend  %>

			</select><br />
			<br />
&nbsp;</td>
<% session("LoginID") = strLoginID
   session("FacultyID") = numFacultyID%> 
	</form>
	</tr>

<form method="post" action="AdminRDateModified.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
<tr>
<th>Select a Facility:</th>
<td>
<%
			 strLoginID = request.form("cboSupervisor")
			 'AA jan 2010 rework for relationship fix
			 strSQL = "Select * from tblfacility, tblFacilitySupervisor where tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID"_
			 &" and strLoginID = '"& strLoginID &"'"
  set rsSearchFacility2 = server.CreateObject("ADODB.Recordset")
  rsSearchFacility2.Open strSQL, Conn, 3, 3 %>
			<select size="1" name="cboFacility">
			<option value="0">Select any one</option>
			<%while not rsSearchFacility2.EOF    %>
			<!-- AA jan 2010 remove lokup by index-->
			<option value="<%=rsSearchFacility2("numFacilityID")%>"><%=cstr(rsSearchFacility2("strRoomName"))+" / "+cstr(rsSearchFacility2("strRoomNumber")) %></option>
			<%rsSearchFacility2.MoveNext 
			wend  %>
			</select>
&nbsp;&nbsp;
	<input type="submit" value="Generate Report" name="btnGenRep" />
	<input type="hidden" name="hdnLoginId" value="<%=strLoginID%>" />
	<input type="hidden" name="hdnFacultyID" value="<%=numFacultyID%>" />
</form>
</td>
</tr>
</table>

</center>

</body>
</html>