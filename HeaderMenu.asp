
<link rel="SHORTCUT ICON" href="favicon.ico" type="image/x-icon" />

<!--img src="images/UTS_logo.png" alt="UTS" width="152" height="50" style="border: 1px; vertical-align:top; float:right;"-->




<div id="wrappertop">

<div id="logo">
<img src="images/UTS_logo.png" alt="UTS" width="120" height="50" style="border: 3px; vertical-align:top; float:left;">
</div>


<div id="content">
<% if session("LoggedIn") then %>
    <% if session("isAdmin") then %>
    <br /> <h1 class="pagetitle">Health and Safety Risk Register V2.4 - Administration Menu</h1>
    <% else %>
     <br /> <h1 class="pagetitle">Health and Safety Risk Register V2.4 - Supervisor Menu</h1>
    <% end if %>
<% else %>
    <br /> <h1 class="pagetitle">Health & Safety Risk Register V2.4</h1>
<% end if %>

<% if session("LoggedIn") then %>
    <div class="topframe">
     You are logged in as <strong><%=session("strName")%></strong><br />
     </div>
<% end if %>

<% if session("isAdmin") then %>
     <div class="loginlist">
     <ul>
    
        <li><a href="Home.asp" title="Search the Online Risk Register">Search</a></li>
  
         <li></li>
        <li><a href="admin.asp" title="Perform administration on the Online Risk Register">Administration Functions</a></li>
        <li><a href="logout.asp" title="Log out of the Online Risk Register">Logout</a></li>
     </ul>
    </div>
<% else %>
    <div class="loginlist">
	 <ul>

	   <li><a href="Home.asp" title="View the Risk Assessments for your facility/facilities">Search</a></li>
	   <!--li><a target="Operation" href="help.htm">Help</a></li-->
	   <!--<li><a target="_top" href="menu.asp" title="Online Risk Register homepage">Home</a></li>-->
         <% if session("LoggedIn") then %>
	        <li><a href="logout.asp" title="Log out of the Risk Register">Logout</a></li>
	          <li><a href="EditSupervisorProfile.asp" title="Edit Profile">My Profile</a></li>
         <% else %>
            <li><a href="#" data-toggle="modal" data-target="#LoginModal" >Login</a></li>
            <!--<li><a href="Register.asp">Create Account</a></li>-->
         <% end if %>
	 </ul>
	</div>
<% end if %>


</div><!-- close the content DIV -->
</div><!-- close the wrapper div -->

<div class="modal fade" id="LoginModal" tabindex="-1" role="dialog" aria-labelledby="LoginModalLabel">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
          <h4 class="modal-title" id="exampleModalLabel">Log In</h4>
      </div>
      <div class="modal-body">
        <form id="loginForm">
        
          <div class="form-group">
            <label for="txtLoginId" class="control-label">User Name:</label>
            <input name="txtLoginID" class="form-control" id="txtLoginID" maxlength="70" type="text" value="<%=strLoginID%>" size="25" />
            </div>
              <div class="form-group">
            <label for="txtPassword" class="control-label">Password:</label>
            <input name="txtPassword" class="form-control" maxlength="70" type="password" size="25" />
          </div>
        </form>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
        <button type="submit" id="submitLogin" class="btn btn-primary">Log In</button>
      </div>
    </div>
  </div>
</div>
   <script type="text/javascript">
   

       $(function () {
           //twitter bootstrap script
           $("#submitLogin").click(function () {
               $.ajax({
                   type: "POST",
                   url: "AJAXLogin.asp",
                   data: $('#loginForm').serialize(),
                   success: function (data) {
                       var obj = jQuery.parseJSON(data);
                       var res = obj.result;
                       if(res == 1){
                           location.reload();
                       }
                       else{
                           alert("Incorrect username or password.  Please try again");
                       }
                       
                   },
                   error: function () {
                       alert("error - please contact adminstrator");
                   }
               });
           });
       });
       </script>