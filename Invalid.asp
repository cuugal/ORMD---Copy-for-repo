<%@Language = VBscript%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="orr.css" media="all" />
</head>
<body>
<%
'Junk everything in session (also kill frames)
	Response.Expires = 0
	session.Abandon 

'3 being the number of seconds before this script redirects back to the main menu
'Response.AddHeader "Refresh", "3;URL=index.htm"
'Response.Redirect "index.htm"

%>
<div id="orrheading"><h1>Online Risk Register</h1><img src="http://www.uts.edu.au/images/css/utslogo.gif" alt="UTS" width="130" height="29" style="border: none; vertical-align: middle; 	float: right; " /><font size="1" color="gray">V1.3</font></div>
<table class="searchtable">
	<tr><td style="text-align:center">
	<br/>
	<br/>
		<h1> You have been logged out</h1>
		<br/>
		<h5> <a href="index.htm" rel="nofollow" target="_top">Back to Home page</a></h5>
	<br/>
	<br/>
	</td></tr>
</table>
</body>
</html>