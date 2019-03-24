<%@Language = VBscript%>

<!--# INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">


<head>
<title>Health and Safety Risk Register</title>
<link rel="SHORTCUT ICON" href="/images/favicon.ico" type="image/x-icon" />
<link rel="apple-touch-icon" href="/images/apple-touch-icon.png"/>
    <!--#include file="bootstrap.inc"--> 
</head>
    <body>
    

      <div id="wrapper" class="container">
         <div id="content">



<!--%@Language = VBscript%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="orr.css" media="all" />
</head>
<body-->


<!-- In case user is directed her by link to SWMS, but does not have an active session -->


<%
'Junk everything in session (also kill frames)
Response.Expires = 0

'3 being the number of seconds before this script redirects back to the main menu
'Response.AddHeader "Refresh", "3;URL=index.htm"
'Response.Redirect "index.htm"

session.Abandon 
%>

<!--meta http-equiv="refresh" content="3; url=home.asp"-->

<table>
	<tr align='center' width='95%'><td>

		<h3> You have been logged out</h3>
		<br/>
		<h5> <a href="Home.asp" target="_top">Back to Home page</a> and try again.</h5>

	</td></tr>
</table>


         </div>
         <!-- close the content DIV -->
      </div>
      <!-- close the wrapper div -->
    </body>
</html>