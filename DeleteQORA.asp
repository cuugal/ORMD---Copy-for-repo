<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"--><%
If Trim(Session("strLoginId")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
 <link rel="stylesheet" type="text/css" href="orr.css" media="screen" />
<title>Online Risk Register - Delete a Risk Assessment</title>

    <!--#include file="bootstrap.inc"--> 
</head>

    <!--#include file="HeaderMenu.asp" -->
<body>
<%
Function Escape(sString)

'Replace any Cr and Lf with <br />
strReturn = Replace(sString , vbCrLf, "<br />")
strReturn = Replace(strReturn , vbCr , "<br />")
strReturn = Replace(strReturn , vbLf , "<br />") 
Escape = strReturn
End Function


dim val 
'dim conn1
dim conn
dim strSQL
dim rsDeleteQORA

val = cint(request.form("hdnQORAId"))
'response.write(val)

  set conn = Server.CreateObject("ADODB.Connection")
  'conn.open constr
  
  strSQL ="delete from tblQORA where numQORAId ="& val 
  
  set rsDeleteQORA = Server.CreateObject("ADODB.Recordset")
 ' rsDeleteQORA.Open strSQL, conn, 3, 3 
  
  conn.open constr
  conn.BeginTrans
  conn.Execute strSQL
  conn.commitTrans
  
  strSQL ="delete from tblRiskControls where numQORAID ="& val 
  
  set rsDeleteQORA = Server.CreateObject("ADODB.Recordset")
 ' rsDeleteQORA.Open strSQL, conn, 3, 3 
  
  'conn.open constr
  conn.BeginTrans
  conn.Execute strSQL
  conn.commitTrans
  
'else
%>
    <div id="wrapper">
  <div id="content">
    <h2 class="pagetitle">Risk Assessment <%=val%> has been deleted successfully</h2>
	
    <div class="addAnother">
 <%
            dim action
            if Session("mostRecentSearch") <> "" then   
                action = Session("mostRecentSearch")
            else
                action = "Home.asp"
            end if
            %>

          <form id="refreshResults" action="<%=action %>" method="post">
            <input type="hidden" name="confirmationMsg" value="" />
            <input type="hidden" name="searchType" value="<%=session("searchType") %>" />
            <input type="hidden" name="cboOperation" value="<%=session("cboOperation")  %>" />
            <input type="hidden" name="cboFacility" value="<%=session("cboFacility") %>" />
            <input type="hidden" name="hdnFacultyId" value="<%=session("cboFaculty") %>" />
            <input type="hidden" name="hdnBuildingId" value="<%=session("hdnBuildingId") %>" />
            <input type="hidden" name="hdnCampusId" value="<%=session("hdnCampusId") %>" />
            <input type="hidden" name="txtHazardoustask" value="<%=session("hdnHTask") %>" />
            <input type="hidden" name="cboSupervisorName" value="<%=session("cboSupervisorName") %>" />
             <input type="submit" class="btn btn-primary" value="Next" name="btnAddMore">
        </form>
        </div>

  </div>
</div>
<!--used to include file="reportAfterEdit.asp"-->
</body>
</html>
