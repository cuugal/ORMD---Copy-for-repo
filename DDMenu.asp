<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%dim loginId
loginId = session("strDLoginId")%>
<%'************************ Getting the login information for displaing it on the menu bar****************************
  dim rsSearchLogin
  dim strName
  dim strSQL
  dim Conn
  
  'Database Connectivity Code 
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
  strSQL = "Select strDGivenName,strDSurname "_
  &" from tblfaculty"_
  &" where tblFaculty.strDLogin = '"& loginId &"'" 
  
  set rsSearchLogin = server.CreateObject("ADODB.Recordset")
  rsSearchLogin.Open strSQL, Conn, 3, 3
  strName = cstr(rsSearchLogin(0)) + " " + cstr(rsSearchLogin(1))  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 <meta http-equiv="Content-Language" content="en-au" />
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />

 <title>Online Risk Register - Dean and Director Menu</title>
<script type="text/javascript">
<!--

function ChangeResults(page){
	parent.frames["Results"].location.href = page
	return true
	}
//-->
</script>

<script type="text/javascript">
<!--
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1); 
function printPage(frame, arg) {
  if (frame == window) {
    printThis();
  } else {
    link = arg; // a global variable
     printFrame(frame);
  }
  return false;
}

function printThis() {
  if (pr) { // NS4, IE5
    window.print();
  } else if (da && !mac) { // IE4 (Windows)
    vbPrintPage();
  } else { // other browsers
    alert("Sorry, your browser doesn't support this feature.");
  }
}

function printFrame(frame) {
  if (pr && da) { // IE5
    frame.focus();
    window.print();
    link.focus();
  } else if (pr) { // NS4
    frame.print();
  } else if (da && !mac) { // IE4 (Windows)
    frame.focus();
    setTimeout("vbPrintPage(); link.focus();", 100);
  } else { // other browsers
    alert("Sorry, your browser doesn't support this feature.");
  }
}
if (da && !pr && !mac) with (document) {
  writeln('<OBJECT ID="WB" WIDTH="0" HEIGHT="0" CLASSID="clsid:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>');
  writeln('<' + 'SCRIPT LANGUAGE="VBScript">');  
  writeln('Sub window_onunload');
  writeln('  On Error Resume Next');  
  writeln('  Set WB = nothing');
  writeln('End Sub'); 
  writeln('Sub vbPrintPage');
  writeln('  OLECMDID_PRINT = 6');
  writeln('  OLECMDEXECOPT_DONTPROMPTUSER = 2');
  writeln('  OLECMDEXECOPT_PROMPTUSER = 1');  
  writeln('  On Error Resume Next');
  writeln('  WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER');
  writeln('End Sub');  
  writeln('<' + '/SCRIPT>');
}
// --></script>
<base target="bottom" />
</head>

<body>
<div id="wrapper">
 <div id="content">
<h1 class="pagetitle">Online Risk Register - Dean and Director Menu</h1>
You are logged in as <strong><%=strName%></strong><strong>.</strong>
&nbsp;&nbsp;&nbsp;&nbsp;Report of Risk Assessments requiring action, AND which have yet to be acted on.<br/>
	<div class="loginlist">
	 <ul>

	   <li><a target="Operation" href="DDReport.asp" title="View Risk Assessment Action Status report for your Faculty/Unit.">RA Action Status</a></li>	 	   
	   <!--li><a target="Operation" href="help.htm">Help</a></li-->  
	   <li><a target="_top" href="homepage.asp" title="Online Risk Register homepage">Home</a></li>
	   <li><a target="_top" href="logout.asp" title="Log out of the Online Risk Register">Logout</a></li>
	 </ul>
	</div>
 </div><!-- close the content DIV -->
</div><!-- close the wrapper div -->
</body>
</html>