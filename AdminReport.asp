<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>
<script Language = "javascript">
function FillDetails()
{
 document.a1.submit();
}

function FillD()
{
 document.b1.submit();  
}
function Reload()
{
 //document.a1.cbofaculty.value = 0;
 //document.a1.cboSupervisor.value = 0;
 window.location.href ="AdminReport.asp";

}
function ReloadB()
{
 document.b1.cbofaculty1.value = 0;
 document.b1.cboSupervisor1.value = 0;
 window.location.href ="AdminReport.asp";

}

</script>

<%dim loginId
loginId = session("strLoginId")%>

<head>

<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Supervisor's report</title>
<%
 dim rsSearchFaculty1
 dim rsSearchFaculty2
 dim Conn
 dim strSQL
 
   '*********************Setting up the database connectivity***********
  set Conn = Server.CreateObject("ADODB.Connection")
  Conn.open constr
        
  '------------------------get the facilities for the login ---------------
  strSQL = "Select * "_
  &" from tblfaculty"
    
  set rsSearchFaculty1 = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty1.Open strSQL, Conn, 3, 3   
 
%>
</head>

<body link="#000000" vlink="#000000" alink="#000000">

<p align="center"><font face="Tahoma"><b>Please select a report </b>
</p>
<div align="center">
	<table border="0" width="70%" id="table1" bordercolor="#FFFFFF" height="226">
		<tr>
			<td width="64" align="center" height="36"><b>
			<font color="#800000" face="Tahoma" size="2">Number</b></td>
			<td height="36" align="center"><b>
			<font color="#800000" face="Tahoma" size="2">Name</b></td>
		</tr>
		<form method="POST" name="a1" action="AdminReport.asp">
			
		<tr>
				
			<td width="64" align="center" rowspan="3"><b>
			1</b></td>
			<td height="19" bordercolor="#FFFFFF" align="left" valign="top">
			<b><font face="Tahoma"><span style="text-decoration: none">
			<font color="#000000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<i>The Risk Level Report<br>
			</i></span>
			<br>
			Select Faculty / Unit&nbsp;
			<%  numfacultyID = cint(request.form("cboFaculty"))
				if numFacultyID = "" then
				numFacultyID = 0
				end if %>
			
			<select size="1" name="cbofaculty" onchange = "javascript:FillDetails()">
			<option value="0">Select All</option>
			  <% if numFacultyID = 0 then
		  response.Write "select All"
		  end if %></option>
        <%while not rsSearchFaculty1.Eof%>
        <option value="<%=rsSearchFaculty1("numFacultyID")%>"
        <% if rsSearchFaculty1("numFacultyID") = numFacultyID then
		  response.Write "selected"
		  end if %>><%=rsSearchFaculty1("strFacultyName")%></option>
        <%rsSearchFaculty1.Movenext
         wend  %>
			</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</b></td>
			
		</tr>
		
			<tr>
			<td align="left" valign="top"><b>
			Select Supervisor</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
	<% 
  if numFacultyId <> 0 then	
   strSQL ="Select * from tblFacilitySupervisor where numFacultyId ="& numFacultyID &" order by strGivenName"
  else
  strSQL ="Select * from tblFacilitySupervisor where strAccessLevel <>'A' order by strGivenName"
  end if 
   set rsFillSup= Server.CreateObject("ADODB.Recordset")
   rsFillSup.Open strSQL, conn, 3, 3
   

     
  strLoginID = request.form("cboSupervisor")
    if strloginId = "" then
	   strLoginId = ""
    end if %>
			<select size="1" name="cboSupervisor" onchange = "javascript:FillDetails()">
			<option value="0"><% if strLoginId = "" then
		 ' response.Write "Select All"
		  end if %>Select All</option>
        <%while not rsFillSup.Eof
           strSurname = rsFillSup("strSurname")
   strGivenName = rsFillSup("strGivenName")
   strName = cstr(strGivenName) + " " + cstr(strSurName)  
   %>
        <option value="<%=rsFillSup("strLoginId")%>"
        <% if rsFillSup("strLoginId") = strLoginID then
		  response.Write "selected"
		  end if %>><%=strName%></option>
        <%rsFillSup.Movenext
         wend 
     
          %>
			</select></td>
		</tr>
		
		</form>
		
		
		<form method="POST" name="a2" action="AdminRLevel.asp">
		<input type = "Hidden" name = hdnFaculty value =<%=numFacultyID%>>
		<input type = "Hidden" name = hdnSupervisor value =<%=strLoginId%>>
		<%
		 if strLoginId <> "" and numfacultyId <> 0 then	
		 'AA jan 2010 fix relationship altered this line
		 '&" where tblfacility.strFacilitySupervisor = tblFacilitySupervisor.strLoginId and"_
         strSQL ="Select * from tblFacility,tblFacilitySupervisor"_
                &" where tblfacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and"_
                &" strLoginID = '"& strLoginId &"' and "_
                &" tblFacilitySupervisor.numFacultyId = "& numFacultyId &" order by strRoomName"
		 else
		  strSQL ="Select * from tblFacility order by strRoomName"
  		end if 
        
        'AA jan 2010 fixing the relationships
       'if strLoginId <> ""  then
        ' strSQL ="Select * from tblFacility where strFacilitySupervisor = '"& strLoginId &"' order by strRoomName"
        'end if
        if strLoginId <> ""  then
        strSQL ="Select * from tblFacility,tblFacilitySupervisor"_
                &" where tblfacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and"_
                &" strLoginID = '"& strLoginId &"' order by strRoomName"
        end if
        
		 set rsFillFaci= Server.CreateObject("ADODB.Recordset")
		 rsFillFaci.Open strSQL, conn, 3, 3

		%>
		<tr>
			<td align="left" valign="top"><b>
			Select Facility&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</b>
			<select size="1" name="cbofacility">
			<option value="0">Select All</option>
			<%while not rsFillFaci.EOF %>
			<option value = <%=rsFillfaci("numFacilityId")%>><%=cstr(rsFillFaci(1)) + " / " + cstr(rsFillFaci(2))%></option>
			<% rsFillFaci.Movenext
			   wend 
			%>
			</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
			<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
			<input type="submit" value="Generate Report" name="btnGenReport">&nbsp;
			<input type="button" value="Clear Form" name="btnClear" onclick = Reload()></td>
		</tr>
		</form>

<%'**********************************************************************************************************************************************************	%>

		<form method="POST" name="b1" action="AdminReport.asp">
			<tr>
<%
  strSQL = "Select * "_
  &" from tblfaculty"
    
  set rsSearchFaculty2 = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty2.Open strSQL, Conn, 3, 3   
 
%>
				
			<td width="64" align="center" rowspan="3"><b>
			2</b></td>
			<td align="left" valign="top"><br>
			<b>
			<i>
			<font face="Tahoma"><span style="text-decoration: none">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; The 
			Renewal Date Report</span><br>
			</i>
			
			<br>
				<%  numfacultyID1 = cint(request.form("cboFaculty1"))
				if numFacultyID1 = "" then
				numFacultyID1 = 0
				end if %>
			Select Faculty / Unit&nbsp;<select size="1" name="cbofaculty1" onchange = "javascript:FillD()">
			<option value="0">Select All</option>
			  <% if numFacultyID1 = 0 then
		  response.Write "select All"
		  end if %></option>
        <%while not rsSearchFaculty2.Eof%>
        <option value="<%=rsSearchFaculty2("numFacultyID")%>"
        <% if rsSearchFaculty2("numFacultyID") = numFacultyID1 then
		  response.Write "selected"
		  end if %>><%=rsSearchFaculty2("strFacultyName")%></option>
        <%rsSearchFaculty2.Movenext
         wend  %>
			</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</b></td>
			
		</tr>
		<%  if numFacultyId1 <> 0 then	
   strSQL ="Select * from tblFacilitySupervisor where numFacultyId ="& numFacultyID1 &" order by strGivenName"
  else
  strSQL ="Select * from tblFacilitySupervisor where strAccessLevel <>'A' order by strGivenName"
  end if 
   set rsFillSup1= Server.CreateObject("ADODB.Recordset")
   rsFillSup1.Open strSQL, conn, 3, 3
   
  
  strLoginID1 = request.form("cboSupervisor1")
    if strloginId1 = "" then
	   strLoginId1 = ""
    end if %>
			<tr>
			<td align="left" valign="top"><b>
			Select Supervisor</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<select size="1" name="cboSupervisor1" onchange = "javascript:FillD()">
			<option value="0"><% if strLoginId1 = "" then
		 ' response.Write "Select All"
		  end if %>Select All</option>
        <%while not rsFillSup1.Eof 
         strSurname = rsFillSup1("strSurname")
   strGivenName = rsFillSup1("strGivenName")
   strName = cstr(strGivenName) + " " + cstr(strSurName)  

        %>
        <option value="<%=rsFillSup1("strLoginId")%>"
        <% if rsFillSup1("strLoginId") = strLoginID1 then
		  response.Write "selected"
		  end if %>><%=strName%></option>
        <%rsFillSup1.Movenext
         wend 
     
          %>
			</select></td>
			</tr>
			</form>
			
		<form method="POST" name="b2" action="AdminRDate.asp">
		<input type = "Hidden" name = hdnFaculty value =<%=numFacultyID1%>>
		<input type = "Hidden" name = hdnSupervisor value =<%=strLoginId1%>>
				<%
		 if strLoginId1 <> "" and numfacultyId1 <> 0 then	
		 'AA jan 2010 fix relationship: altered this line
		 '&" where tblfacility.strFacilitySupervisor = tblFacilitySupervisor.strLoginId and"_
         strSQL ="Select * from tblFacility,tblFacilitySupervisor"_
                &" where tblfacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and"_
                &" strLoginID = '"& strLoginId1 &"' and "_
                &" tblFacilitySupervisor.numFacultyId = "& numFacultyId1 &" order by strRoomName"
		 else
		  strSQL ="Select * from tblFacility order by strRoomName"
  		end if 
        
        'AA jan 2010 relationship fix
        'if strLoginId1 <> ""  then
        ' strSQL ="Select * from tblFacility where strFacilitySupervisor = '"& strLoginId1 &"' order by strRoomName"
        'end if
        if strLoginId <> ""  then
        strSQL ="Select * from tblFacility,tblFacilitySupervisor"_
                &" where tblfacility.numFacilitySupervisorID = tblFacilitySupervisor.numSupervisorID and"_
                &" strLoginID = '"& strLoginId &"' order by strRoomName"
        end if
  
		 set rsFillFaci1= Server.CreateObject("ADODB.Recordset")
		 rsFillFaci1.Open strSQL, conn, 3, 3

		%>
		
			<tr>
			<td align="left" valign="top"><b>
			Select Facility&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</b><select size="1" name="cbofacility1">
			<option value="0">Select All</option>
			<%while not rsFillFaci1.EOF %>
			<option value = <%=rsFillfaci1("numFacilityId")%>><%=cstr(rsFillFaci1(1)) + " / " + cstr(rsFillFaci1(2))%></option>
			<% rsFillFaci1.Movenext
			   wend 
			%>
			</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
			<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
			<input type="submit" value="Generate Report" name="btnGenReport0">&nbsp;
			<input type="button" value="Clear Form" name="btnClear" onclick = Reload() ></td>
			</tr>
		</form>
		<tr>
				
			<td width="64" align="center">&nbsp;</td>
			<td align="left" valign="top">&nbsp;</td>
		</tr>
	</table>
</div>

</body>

</html>