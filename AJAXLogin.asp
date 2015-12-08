<!--#include file="aspJSON.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->
<%
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
    Set oJSON = New aspJSON
    With oJSON.data
    .Add "result", -1

    End With
    Response.Write oJSON.JSONoutput()  
   
elseIf  rsAccess2("strPassword")= strPassword then
	session("LoggedIn")= true
	session("strLoginID")= strLoginID
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
    session("strAccessLevel")  = rsAccess2("strAccessLevel")

    if rsAccess2("strAccessLevel") ="A" then
		session("isAdmin") = true
	elseif rsAccess2("strAccessLevel") ="S" then
		session("isAdmin") = false
	end if 
		
    
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

    Set oJSON = New aspJSON
    With oJSON.data
    .Add "result", 1

    End With
    Response.Write oJSON.JSONoutput()  

else
    Set oJSON = New aspJSON
    With oJSON.data
    .Add "result", -1

    End With
    Response.Write oJSON.JSONoutput()  
end if




%>
