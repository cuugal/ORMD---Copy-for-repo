<!--#INCLUDE FILE="DbConfig.asp"-->
<html>

<head>
<title>Search Results For UTS Risk Assessments</title>
<!--#include file="bootstrap.inc"--> 
</head>
    <body>
<!--include file="ReportSearchHeader.asp" -->
        <!--#include file="HeaderMenu.asp" -->
          <%
        dim confirmationMsg
        confirmationMsg = Session("confirmationMsg")
        if confirmationMsg <> "" then
         %>
            <div class="wrapper">
              <div class="content">
                <h1 class="pagetitle"><%=confirmationMsg%> </h1>
              </div>
            </div>
    <% end if %>

<!--#include file="searchResults.asp"--> 


        </body>
</html>
