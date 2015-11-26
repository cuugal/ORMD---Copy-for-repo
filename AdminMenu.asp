 
<%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>

<%dim loginId
loginId = session("strLoginId")%>

<%'******* Getting the login information for display on the menu bar *****

    dim strNameLogin
    strNameLogin = session("strName")

%>

<script type="text/javascript">
<!--

function ChangeResults(page){
	parent.frames["Results"].location.href = page
	return true
	}
//-->

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
// -->
</script>



<div id="wrappertop">
<div id="content">

<h1 class="pagetitle">Online Risk Register - Administration Menu</h1>
<div class="topframe">
 You are logged in as <strong><%=strNameLogin%></strong><br />
 </div>
 <div class="loginlist">
 <ul>
 <li><a target="Operation" href="LocationAdmin.asp" title="Create a new Risk Assessment">Create Risk Assessment</a></li>
  <li><a target="Operation" href="searchQORA.asp" title="Search the Online Risk Register">Search Risk Assessments</a></li>
  
  <li></li>
  <li><a target="Operation" href="admin.asp" title="Perform administration on the Online Risk Register">Administration Functions</a></li>
 <!-- <li><a target="Operation" href="MyQoraAdmin.asp" title="&lsquo;My RAs&rsquo;">My RAs</a></li>
  <li><a target="Operation" href="help.htm" title="View the documentation for the Online Risk Register">Help</a></li>
  <li><a target="_top" href="menu.asp" title="Go to the home page of the Online Risk Register">Home</a></li>-->
  <li><a target="_top" href="logout.asp" title="Log out of the Online Risk Register">Logout</a></li>
 </ul>
</div>

</div><!-- close the content DIV -->
</div><!-- close the wrapper div -->
