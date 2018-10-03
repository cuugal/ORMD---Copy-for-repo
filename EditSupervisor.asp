<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
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
Dim rsFillLoginId
Dim strSQL
Dim numSupervisorID

'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor where strAccessLevel <> 'A' order by strLoginId"
   set rsFillLoginId = Server.CreateObject("ADODB.Recordset")
   rsFillLoginId.Open strSQL, conn, 3, 3
%>
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <!--#include file="bootstrap.inc"--> 
<title>Online Risk Register - Edit a User</title>
<script type="text/javascript">
// function to ask about the confirmation of the file. 
function ConfirmChoice() 
{
 if($("#strAccessLevel").val() == 'T' && parseInt($("#counts").val()) > 0){
       alert ("Please unassign this user from any operations or facilities before changing to Assessor");
     	   return(false);
 }


  if ((document.Form1.cboFaculty.value != "0") && (document.Form1.txtSurname.value != "") && (document.Form1.txtGivenName.value !="") && (document.EditSupervisor.cboLoginId.value !="0") && (document.Form1.txtPassword.value !="") && (document.Form1.txtnewID.value != "")) 
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
	
      alert ("Any fields on the form cant be empty , please fill in the entire form !");
	   return(false);
	}
}

// function to reload the form to add the new entries
function FillDetails()
{
 document.EditSupervisor.submit();

}
</script>
<script type="text/VBScript">
  function clearform()
  dim val 
  val = ""
  document.EditSupervisor.cboLoginId.value = 0
  document.Form1.cboFaculty.value = 0
  document.Form1.txtGivenName.value  = val
  document.Form1.txtSurname.value  = val
  document.Form1.txtPassword.value  = val
  
  end function
</script>
</head>

<body>
    <!--#include file="HeaderMenu.asp" -->
<div id="wrapper">
 <div id="content" class="contentcenter">

 <h2 class="pagetitle">Edit a User</h2>
 


 <table class="adminfn" >
 <tr>
  <th>Select the Login ID you wish to change:</th>
  <td><form method="post" action="EditSupervisor.asp" name="EditSupervisor">
        <% strLoginID = request.form("cboLoginId")
        	
        if strloginID = "" then
	       strLoginID = "0"
        end if %>
      <select size="1" name="cboLoginId" onchange="javascript:FillDetails()">
      <option value="0" 
        <% if strLoginID = "0" then
		  response.Write "select any one"
		  end if %>>Select any one</option>
        <% while not rsFillLoginId.Eof%>
        <option value="<%=rsFillLoginId("strLoginID")%>"
        <% if rsFillLoginId("strLoginID") = strLoginID then
		  response.Write "selected='selected'"
		  end if %>><%=rsFillLoginId("strLoginId")%></option>
        <%rsFillLoginId.Movenext
         wend 
         
         ' closing the connections
         
           rsFillLoginId.close
           set rsFillLoginId = nothing
           conn.Close
           set conn = nothing
         %>
      </select><input type="hidden" name="hdnLoginId" value="<%=strLoginId%>" />
      </form>
      </td>
    </tr>
    <tr>
    <form method="post" action="AdminEdit.asp" name="Form1" enctype="application/x-www-form-urlencoded" onsubmit="return ConfirmChoice();">
    <input type="hidden" name="hdnLoginId" value="<%=strLoginId%>" />
   
<%
Dim connDet
Dim rsFillDetails


'Database Connectivity Code 
  set connDet = Server.CreateObject("ADODB.Connection")
  connDet.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFacilitySupervisor where strLoginId = '"& strLoginId &"'"
   set rsFillDetails = Server.CreateObject("ADODB.Recordset")
   rsFillDetails.Open strSQL, connDet, 3, 3
   
  
%>

<tr>
<th>User Type:</th>
<td>
 <% if rsFillDetails.EOF <> true then%>
<select name="strAccessLevel" id="strAccessLevel" value="<%=rsFilldetails("strAccessLevel") %>">
<option value="S" <% if rsFillDetails("strAccessLevel") = "S" then %>selected<% end if %>>Supervisor</option>
<option value="T"  <% if rsFillDetails("strAccessLevel") = "T" then %>selected<% end if %>>Assessor</option>
</select>
<% else %>
<select size="1" name="strAccessLevel" >
      <option>No Records</option>
      </select>
<% end if %>
</td>
</tr>


	<tr>
      <th>Proposed new Login ID:</th>
      <td>
       <input type="text" name="txtnewID" size="20" tabindex="9" value="<%= strLoginID%>" />
         
      </td>
    </tr>

    <tr>
      <th>User's Surname:</th>
      <td>
      <% if rsFillDetails.EOF <> true then%>


      <input type="text" name="txtSurname" size="20" tabindex="1" value="<%= rsFilldetails("strSurname")%>" />
      <input type="hidden" name="hdnSupervisorId" value="<%=rsFillDetails("numSupervisorID")%>" /></td>
      </tr>

      <tr id="userfaculty">
      <th >User's Faculty/Unit:</th>
      <td>
     <%'code to add the related faculties for that str loginID 
      Dim connFac
      Dim rsFillFac
     
     'Database Connectivity Code 
      set connFac = Server.CreateObject("ADODB.Connection")
      connFac.open constr
 
     ' setting up the recordset
       strSQL ="SELECT * FROM tblFacilitySupervisor, tblFaculty WHERE tblFacilitySupervisor.numFacultyId = tblFaculty.numFacultyID and strLoginId = '"& strLoginId &"'"
       set rsFillFac = Server.CreateObject("ADODB.Recordset")
       rsFillFac.Open strSQL, connFac, 3, 3

     'code to add the different faculties for that str loginID
      Dim connFaculty
      Dim rsFillFaculty
     
     'Database Connectivity Code 
      set connFaculty = Server.CreateObject("ADODB.Connection")
      connFaculty.open constr
 
     ' setting up the recordset
       strSQL ="Select * from tblFaculty order by strFacultyName"
       set rsFillFaculty = Server.CreateObject("ADODB.Recordset")
       rsFillFaculty.Open strSQL, connFac, 3, 3

       'check to see if this supervisor is assigned any operations or facilities

         ' setting up the recordset
          strSQL ="SELECT * FROM tblFacilitySupervisor WHERE strLoginId = '"& strLoginId &"'"
          set rsFillSup = Server.CreateObject("ADODB.Recordset")
          rsFillSup.Open strSQL, connFac, 3, 3


       strSQL = "SELECT count( *) as cc from tblOperations, tblFacility where tblOperations.NumfacilitySupervisorId = "&rsFillSup("numSupervisorId")&" or tblFacility.numfacilitySupervisorId = "&rsFillSup("numSupervisorId")&""
        set rsCount = Server.CreateObject("ADODB.Recordset")
        rsCount.Open strSQL, connFac, 3, 3


      %>

      <input type="hidden" value="<%=rsCount("cc") %>" id="counts"/>

      <select size="1" id="cboFaculty" name="cboFaculty" tabindex="3">
   <% if rsFillDetails("strAccessLevel") = "S" then %>
      <option value="<%=rsFillFac("tblFacilitySupervisor.numFacultyID")%>" selected="selected"><%=rsFillFac("strFacultyName")%></option>
          <% end if %>
      <%While not rsFillFaculty.EOF%>
      
      <option value="<%=rsFillFaculty("numFacultyID") %>"><%=rsFillFaculty("strFacultyName")%></option>
      
      <%
        rsFillFaculty.MoveNext
        wend %> 
      </select></td>
    </tr>

    <tr>
      <th>User's Given Name:</th>
      <td>
       <input type="text" name="txtGivenName" size="20" tabindex="2" value="<%= rsFillDetails("strGivenName")%>" />
      </td>
    </tr>
    
    <tr>
      <th>User's Password:</th>
      <td><input type="password" name="txtPassword" size="20" tabindex="4" value="<%= rsFillDetails("strPassword")%>" /></td>
    </tr>
    <tr>
      <th>Active:</th>
      <td align="center"><input type="checkbox" id="deprecated" name="deprecated" tabindex="5" <%if not rsFillDetails("boolDeprecated") then%>checked <%end if%>/></td>
    </tr>

<!--end of case 1-->
    <%else%>

 <input type="text" name="txtSurname" size="20" tabindex="1" value="" />
 </td>
</tr>

<tr id="userfaculty">
 <th>User's Faculty/Unit:</th>
 <td>
      <select size="1" id="cboFaculty" name="cboFaculty" tabindex="3">
      <option>No Records</option>
      </select>
  </td>
</tr>

<tr>
 <th>User's Given Name:</th>
 <td>
   <input type="text" name="txtGivenName" size="20" tabindex="2" value="" />
 </td>
</tr>

<tr>
 <th>User's Password:</th>
 <td>
   <input type="password" name="txtPassword" size="20" tabindex="4" value="" />
 </td>
</tr>
<tr>
      <th>Active:</th>
      <td align="center"><input type="checkbox" name="deprecated" tabindex="5" /></td>
    </tr>
<!--end of case 2-->
 <%end if%>

<tr>
 <td colspan="2">

 <input type="hidden" name="hdnOption" value="Supervisor" />  
 <input type="submit" value="Edit" name="btnSave" tabindex="5" /> &nbsp;
 <!-- CL note: This button does not work in Mozilla etc. as it calls VBscript - replaced with a reset button instead 
 <input type="Button" value="Clear" name="btnClear" tabindex="6" onclick = clearform()>
  -->
 <input type="reset" value="Reset" name="btnClear" />
</form>
 </td>
</tr>
</table>
<script type="text/javascript">

$("#strAccessLevel").change(function(){

    if($(this).val() == 'S'){
        $("#userfaculty").show();

    }
    else{
        $("#userfaculty").hide();
    }
});

$(document).ready(function(){

    if($("#strAccessLevel").val() == 'S'){
            $("#userfaculty").show();
        }
        else{
            $("#userfaculty").hide();
        }
});
</script>


</div></div>

</body>
</html>