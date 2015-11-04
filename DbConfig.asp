<%
Dim constr
constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("Database/Ormd.mdb")

Function InjectionEncode(str)
	InjectionEncode=Replace(str,"'","''")
End Function

//one week session expiry
Session.Timeout=1440

If Trim(Session("My_Session")) = "" Then
Response.Redirect("Invalid.asp")
End If
%>
