<%@Language = VBscript%>
<% session("My_Session") = "Open" %>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">

<%
'Set this session variable to prevent displaying of dates in american format (Access has no proviso for normal dates, only US)
session.LCID = 2057	'English(British) format

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



//This focuses the user client on the login ID form, for ease of login - CL 9/5/2011
function formfocus() {
   document.getElementById('txtLoginID').focus();
   }
 window.onload = formfocus;
</script>
<head>
    <!--#include file="bootstrap.inc"--> 
 <title>Online Risk Register - Search health and safety risk assessments</title>
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
  		strSQL = "Select strGivenName,strSurname, numFacultyId, numSupervisorId "_
  		&" from tblfacilitySupervisor"_
  		&" where tblFacilitySupervisor.strLoginId = '"& strLoginID &"'" 
  
  		set rsSearchLogin = server.CreateObject("ADODB.Recordset")
  		rsSearchLogin.Open strSQL, Conn3, 3, 3
  		
  		strName = cstr(rsSearchLogin(0)) + " " + cstr(rsSearchLogin(1))
  		session("strName") = strName
		session("numSupervisorId") = rsSearchLogin("numSupervisorId")
        session("numFacultyId") = rsSearchLogin("numFacultyId")
		
    
        dim strFacultyName
        strFacultyName = "-"
        if rsSearchLogin("numFacultyId") <> -1 then
            strSQL = "Select strFacultyName "_
  		    &" from tblfaculty"_
  		    &" where numFacultyId = "& rsSearchLogin("numFacultyId")
  
  		    set rsSearchLogin = server.CreateObject("ADODB.Recordset")
  		    rsSearchLogin.Open strSQL, Conn3, 3, 3
            strFacultyName = rsSearchLogin("strFacultyName")
        end if  

	    session("strFacultyName") = strFacultyName	



        if rsAccess2("strAccessLevel") ="A" then
		   Response.Redirect "indexLoggedAdmin.asp"
		elseif rsAccess2("strAccessLevel") ="S" then
		   Response.Redirect "indexLoggedSupervisor.asp"
		   elseif rsAccess2("strAccessLevel") ="D" then
		   Response.Redirect "indexLoggedDD.asp"
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
    
    <!--#include file="searchQORA.asp"--> 

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