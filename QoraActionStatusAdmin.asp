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

<head>
<script type="text/javascript">
// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  if ((document.Form1.cboFacultyName.value == "0") ) 
  {
      alert ("Please select a Faculty/Unit to proceed!");
	   return(false);
	}
  else if ((document.Form1.cboSupervisor.value != "0") && (document.Form2.cboFacility.value =="0"))
  {
      alert ("Please select both a supervisor and their respective facility.");
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
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - RA Action Status Report - Administrator</title>
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

<table class="myqora" style="width: 75%;">
<thead>
<tr>
 <th colspan="2">RA Action Status</th>
</tr>
</thead>
<tbody>
<tr>
 <td colspan="2">This report shows UTS Risk Assessments for specified faculties/units where risk controls have not yet been implemented.</td>
</tr>

<tr>
<form method="post" action="QoraActionStatusAdmin.asp" name="Form1">
  <%strSQL = "Select * from tblfaculty"
  
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty.Open strSQL, Conn, 3, 3 %>
				
<th>
      <% numFacultyID = cint(request.form("cboFacultyName"))
      if numFacultyID = "" then
       numFacultyID = 0
        end if %>
Faculty/Unit:</th>
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
			<tr>
<!-- Dean/Director extraction code not working - noted by CL 26-6-2006

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			Faculty's Dean/Director Name : <%
			if rsSearchFacultyDD.EOF <> True then
			 response.write(rsSearchFacultyDD("strDGivenName")+" "+rsSearchFacultyDD("strDsurName"))
			end if %> <br>
			
&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; --->
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
			<option value = "<%=rsSearchsup(0)%>" 
			<% if rsSearchsup("strloginID") = strLoginID then
		          response.Write "selected"
		       end if %>><%=cstr(rsSearchsup(3))+"  "+cstr(rsSearchsup(4)) %></option>
			<%rsSearchsup.MoveNext 
			wend  %>

			</select></td>
			</td>
<% session("LoginID") = strLoginID
   session("FacultyID") = numFacultyID%> 
			</form>
		</tr>
<tr>

<form method="post" action="QASARDateModified.asp" name="Form2" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
<th>Select a Facility:</th>
<td>
	 <%
	 strLoginID = request.form("cboSupervisor")
	 'AA jan 2010 fix for relationship
	 strSQL = "Select * from tblfacility, tblFacilitySupervisor where tblFacilitySupervisor.numSupervisorID = tblFacility.numFacilitySupervisorID and strLoginID = '"& strLoginID &"'"
  
  set rsSearchFacility2 = server.CreateObject("ADODB.Recordset")
  rsSearchFacility2.Open strSQL, Conn, 3, 3 %>

	<select size="1" name="cboFacility">
	<option value="0">Select any one</option>
	<%while not rsSearchFacility2.EOF    %>
	<!--aa jan 2010 remove lookup by index rplce with lookup by name-->
	<option value="<%=rsSearchFacility2("numFacilityID")%>"><%=cstr(rsSearchFacility2("strRoomName"))+" / "+cstr(rsSearchFacility2("strRoomNumber")) %></option>
	<%rsSearchFacility2.MoveNext 
	wend  %>
	</select>&nbsp;&nbsp;
	<input type="submit" value="Generate Report" name="btnGenRep" />
	<input type="hidden" name="hdnLoginId" value="<%=strLoginID%>" />
	<input type="hidden" name="hdnFacultyID" value="<%=numFacultyID%>" />
</form>
</td>
</tr>
</tbody>
</table>
</center>
</body>
</html>