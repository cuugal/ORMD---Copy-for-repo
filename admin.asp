<%@Language = VBscript%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="all" />
 <style type="text/css">
 .navcontainer { width: 200px; font-size: 9pt; }

.navcontainer ul { margin-left: 0; padding-left: 0; list-style-type: none; font-family: Arial, Helvetica, sans-serif; }

.navcontainer a { display: block; padding: 3px;  width: 160px;  background-color: #EDF4F6; border: 1px solid #D7E9ED; margin-bottom: 2px;}

.navcontainer a:link, .navlist a:visited { color: #0083B3; text-decoration: none; margin-bottom: 2px;}

.navcontainer a:hover { background-color: #F8E9AD; color: #0083B3; margin-bottom: 2px;}
 </style>
 <title>Online Risk Register - Administration Functions</title>
</head>
<body>
    <!--#include file="adminMenu.asp" -->
<%
 if session("strLoginId") <> "admin" then
  response.redirect "AccessRestricted.htm"
 end if
%>
<div id="wrapper">
<div id="content">

<h1 class="pagetitle">Administration Functions</h1>

<table>
<tr>
	<td>
   <div class="navcontainer">
   <ul class="navlist">
 	<li><a target="_self" href="CreateFaculty.asp" title="Create a Faculty/Unit">Create a Faculty/Unit</a>
	<li><a target="_self" href="CreateCampus.asp" title="Create a Campus">Create a Campus</a></li>
	<li><a target="_self" href="CreateBuilding.asp" title="Create a Building">Create a Building</a></li>
	<li><a target="_self" href="createSupervisor.asp" title="Create a Supervisor">Create a Supervisor</a></li>
	<li><a target="_self" href="CreateFacility.asp" title="Create a Facility">Create a Facility</a></li>
	<li><a target="_self" href="CreateOperation.asp" title="Create an Operation">Create an Operation</a></li>
</ul>
</div></td>
	<td>
  <div class="navcontainer">
   <ul class="navlist">
     	<li><a target="_self" href="EditFaculty.asp" title="Edit a Faculty/Unit">Edit a Faculty/Unit</a></li>
	<li><a target="_self" href="EditCampus.asp" title="Edit a Campus">Edit a Campus</a></li>
	<li><a target="_self" href="EditBuilding.asp" title="Edit a Building">Edit a Building</a></li>
	<li><a target="_self" href="EditSupervisor.asp" title="Edit a Supervisor">Edit a Supervisor</a></li>
	<li><a target="_self" href="EditFacility.asp" title="Edit a Facility">Edit a Facility</a></li>
	<li><a target="_self" href="EditOperation.asp" title="Edit an Operation">Edit an Operation</a></li>
   </ul>
  </div>
  </td>
	<td>
    <div class="navcontainer">
   <ul class="navlist">
	<li><a target="_self" href="DeleteFaculty.asp" title="Delete a Faculty/Unit">Delete a Faculty/Unit</a></li>
	<li><a target="_self" href="DeleteCampus.asp" title="Delete a Campus">Delete a Campus</a></li>
	<li><a target="_self" href="DeleteBuilding.asp" title="Delete a Building">Delete a Building</a></li>
	<li><a target="_self" href="DeleteSupervisor.asp" title="Delete a Supervisor">Delete a Supervisor</a></li>
	<li><a target="_self" href="DeleteFacility.asp" title="Delete a Facility">Delete a Facility</a></li>
	<li><a target="_self" href="DeleteOperation.asp" title="Delete an Operation">Delete an Operation</a></li>
   </ul>
  </div>
  </td>
</tr>
</table>

</body>
</html>