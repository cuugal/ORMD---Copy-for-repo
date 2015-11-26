
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
Response.Expires=0
Response.Buffer = True

If Request.Form("btnDLogin") = "Login" then

Dim strLoginID, strPassword, msg
dim strAccessLevel
Dim strSQL
Dim rsAccess
Dim conn 

strLoginID=request.form("txtLoginID")
strPassword=request.form("txtPassword")
strSQL="select * from tblFaculty where strdLogin='"& strLoginID &"'"



set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsAccess = Server.CreateObject("ADODB.Recordset")
rsAccess.Open strSQL, conn, 3, 3
If  rsAccess.eof then
   'msg = "The login ID: " + strLoginID + " does not exist, try a different Login ID. Please contact administrator if you need a login ID and Password"
   'Response.Write(msg) %> 
 <script type="text/javascript">
   alert("Invalid username or password - please try again.")
 </script>
   
<%else If  rsAccess("strDPassword")= strPassword then
		session("DLoggedIn")= true
		session("strDLoginID")= strLoginID
		'strAccessLevel = rsAccess("strAccessLevel")
		'session("strAccessLevel") = strAccessLevel
		
		'if rsAccess("strAccessLevel") ="A" then
		 '  Response.Redirect "indexLoggedAdmin.htm"
		'elseif rsAccess("strAccessLevel") ="S" then
		 '  Response.Redirect "indexLoggedSupervisor.htm"
		  ' elseif rsAccess("strAccessLevel") ="D" then
		   Response.Redirect "IndexDD.htm"
		'end if   
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
		 <script type="text/javascript">
			alert("Your session has timed out - please re-login.")
		</script>
   
		<%
	end if
End if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <title>Online Risk Register - Dean and Director Login</title>

<body>
<div id="wrapper">
<div id="content">

<h1 class="pagetitle">Dean and Director Login</h1>

<center>
<form action="loginDD.asp" name="frmlogin" method="post">
<table class="bluebox" style="width: 40%;">
<tr>
   <th>Login ID:</th>
   <td><input name="txtLoginID" maxlength="50" type="text" value="<%=strLoginID%>" size="20" /></td>
</tr>
<tr>
   <th>Password:</th>
   <td><input name="txtPassword" maxlength="50" type="password" size="20" /></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" value="Login" name="btnDLogin" />&nbsp;&nbsp;&nbsp; <input type="reset" value="Clear" name="btnReset" /></td>
</tr>
<TR>
	<td colspan ="2">NOTE: Login and password are case sensitive.</td>
</TR>
</table>
</form>

<br/>
<div class="loginlist">
<ul>
 <li><a target="_self" href="Login.asp" title="The Online Risk Register login for Supervisors and Administrators.">Supervisor and Administrator Login</a></li>
 <!--li><a target="_self" href="help.htm" title="Read the Online Risk Register documentation.">Help</a></li-->
 <li><a target="_top" href="menu.asp" title="Return to the Online Risk Register home page.">Home</a></li>
</ul>
</div>
</center>

</div>

</body>
</html>