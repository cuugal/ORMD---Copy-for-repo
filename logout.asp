<%@ Language=VBScript %>
<%
	Option explicit
	Response.Expires = 0
%>

<%	
	session.Abandon 
	Response.Redirect "index.htm"
%>