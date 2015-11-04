<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />

<%dim loginId
loginId = session("strDLoginId")%>

<script language="javascript">

// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
  //if ((document.Form1.cboSupervisor.value== "0")  )
  //{
  //    alert ("Please select atleast supervisor ");
   //    return(false);
  //} 
//   else if (document.Form2.cboFacility.value =="0")
//  {
//    alert ("Please select both the supervisor and a facility.");
//       return(false);
//  }
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

// Function added by Campbell on 14/6/2006 to clear form contents
function clearform()
{
  var str 
  str = "DDReport.asp";
  window.location.replace(str); 
}
</script>
<head>

<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Online Risk Register - Action Status Login</title>
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
<div id="wrapper">
<div id="content">

<h1 class="pagetitle">Online Risk Register - RA Action Status Report</h1>

<center>

<table class="myqora" style="width: 50%;">
<thead>
<tr>
    <th colspan="2">Risk Assessment Action Status</th>
</tr>
</thead>
<tbody>
<tr>
    <form method="post" action="DDReport.asp" name="Form1">

  <%strSQL = "Select * from tblfaculty where tblFaculty.strDLogin = '"& loginId &"'"
  
  set rsSearchFaculty = server.CreateObject("ADODB.Recordset")
  rsSearchFaculty.Open strSQL, Conn, 3, 3 
  
  numFacultyId = rssearchFaculty("numFacultyID")%>
            <th>Faculty/Unit</th>
            <td><%=rsSearchFaculty("strFacultyName")%></td>
        </tr>
        <tr>
            <th>Dean/Director</th>
            <td><%
            if rsSearchFaculty.EOF <> True then
             response.write(rsSearchFaculty("strDGivenName")+" "+rsSearchFaculty("strDsurName"))
            end if %></td>
        </tr>
        <tr>
            <th>Supervisor Name</th>
            <td><%      
            strLoginID = Request.Form("cboSupervisor") 
            if strLoginID = "" then
           strLoginID = ""
        end if
            
  strSQL = "Select * from tblfacilitySupervisor where numFacultyID = "& numFacultyID
  set rsSearchSup = server.CreateObject("ADODB.Recordset")
  rsSearchSup.Open strSQL, Conn, 3, 3 %>
  
  
            <select size="1" name="cboSupervisor" onChange="javascript:FillFaculty()">
            <option value="0" 
            <% if strLoginId = "" then
                    response.Write "select any one"
             end if %>>Select All</option>
             
            <%while not rsSearchsup.EOF    %>
            <option value="<%=rsSearchsup(0)%>" 
            <% if rsSearchsup("strloginID") = strLoginID then
                  response.Write "selected"
               end if %>><%=cstr(rsSearchsup(3))+"  "+cstr(rsSearchsup(4)) %></option>
            <%rsSearchsup.MoveNext 
            wend  %>
            </select>
            </td>
        </tr>
<% session("LoginID") = strLoginID
   session("FacultyID") = numFacultyID%> 
        </form>
<tr>
<th><form method="post" action="DDRDateModified.asp" name="Form2" enctype="application/x-www-form-urlencoded" onSubmit="return ConfirmChoice();">
Facility</th>&nbsp;
            
             <%
             strLoginID = request.form("cboSupervisor")
             'AA jan 2010 relationship fix - added this to query to get the numSupervisorID
             

             'strSQL = "Select * from tblfacility where strFacilitySupervisor = '"& strLoginID &"'"
             strSQL = "Select * from tblfacility, tblFacilitySupervisor where tblFacilitySupervisor.numSuperVisorID = tblFacility.numFacilitySupervisorID "_
             	&"and strLoginID = '"& strLoginID &"'"
  
  set rsSearchFacility2 = server.CreateObject("ADODB.Recordset")
  rsSearchFacility2.Open strSQL, Conn, 3, 3 %>

            <td><select size="1" name="cboFacility">
            <option value="0">Select All</option>
            <%while not rsSearchFacility2.EOF    %>
            <!-- AA jan 2010 replaced lookup by index with lookup by name-->
            <option value="<%=rsSearchFacility2("numFacilityID")%>"><%=cstr(rsSearchFacility2("strRoomName"))+" / "+cstr(rsSearchFacility2("strRoomNumber")) %></option>
            <%rsSearchFacility2.MoveNext 
            wend  %>
            </select>&nbsp;
            </td>
            </tr>
<tr>
 <td colspan="2">
  <center>
   <input type="submit" value="Generate Report" name="btnGenRep" />&nbsp;&nbsp;&nbsp;
   <!-- clear buttonj added by Campbell on 16/6/2006 -->
   <input type="button" value="Reset" name="btnClear" onclick="clearform()" />
   <input type="hidden" name="hdnLoginId" value="<%=strLoginID%>" />
   <input type="hidden" name="hdnFacultyID" value="<%=numFacultyID%>" /></center>
   </form>
</td>
</tr>
</tbody>
</table>
</div>
</div>

</body>

</html>