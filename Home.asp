<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">


<head>
<title>Online Risk Management Database System</title>
    <!--#include file="bootstrap.inc"--> 
</head>
    <body>

<!--#include file="HeaderMenu.asp" -->
        
      <div id="wrapper" class="container">
         <div id="content">
            <h1 class="pagetitle">Search UTS Risk Assessments</h1>
            <center>
                <div style="width:950px">
                    <h3 style="float:left;color:#6edb04">Search By:</h3>
                </div>
                <div style="clear:both"></div>
                <div style="width:950px">
                    <!--#include file="searchQORA.asp"--> 
                </div>
            </center>
             <br />
              <br />
         </div>
         <!-- close the content DIV -->
      </div>
      <!-- close the wrapper div -->
    </body>
</html>