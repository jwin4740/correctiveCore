<!--#include virtual="/includes/global.asp" -->
<%
   Response.Redirect("login.asp") 
%>
<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">
    <title>User Registration</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">      
    
    <link href="../../scripts/bootstrap/css/bootstrap.min.css" rel="stylesheet">
    <link href="../../scripts/bootstrap/css/bootstrap-responsive.min.css" rel="stylesheet">
    <link href="../../scripts/icons/general/stylesheets/general_foundicons.css" media="screen" rel="stylesheet" type="text/css" />  
    <link href="../../scripts/icons/social/stylesheets/social_foundicons.css" media="screen" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="../../scripts/fontawesome/css/font-awesome.min.css">
    <link href="../../scripts/carousel/style.css" rel="stylesheet" type="text/css" />
    <link href="../../scripts/camera/css/camera.css" rel="stylesheet" type="text/css" />
    <link href="http://fonts.googleapis.com/css?family=Allura" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Aldrich" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Pacifico" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Palatino+Linotype" rel="stylesheet" type="text/css">
    <link href="http://fonts.googleapis.com/css?family=Calligraffitti" rel="stylesheet" type="text/css">

    <link href="../../styles/custom.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="layout/top-index.asp" -->
<br /><br />
<div class="row-fluid">

<form action="jwUserRegistrationProcess.asp" method="post">
<div class="span6">
 <%If Session("jwUserRegistrationErrorMsg") <> "" Then %>
        <h5 style='margin-top: 0px; color: red; font-weight: bolder;' id="errorAlert"><%=Session("jwUserRegistrationErrorMsg")%></h5> 
    <%
        Session("jwUserRegistrationErrorMsg") = ""
    End If 
    %>



<h3>Create username and password</h3>
<input type="text" name="jwRegistrationUsername" id="User Name" value=""  class="input-block-level" placeholder="User Name" />
<input type="password" name="jwRegistrationPassword" id="Password" value=""  class="input-block-level" placeholder="Password" />
<input type="password" name="jwRegistrationConfirmPassword" id="confirmPassword" value=""  class="input-block-level" placeholder="Confirm Password" /> 
<br>
<h3>Program Registered For</h3>
<label class="inline">
<input type="radio" name="programChoice" value="prenatal"> Prenatal <br> 
</label>
<label class="inline">
<input type="radio" name="programChoice" value="foundation"> Foundation <br>
</label>
<br/><br />

<h3>Contact Information</h3>
<input type="text" name="First Name" id="First Name" value=""  class="input-block-level" placeholder="First Name" />
<input type="text" name="Last Name" id="Last Name" value=""  class="input-block-level" placeholder="Last Name" />
<input type="text" name="DOB" id="DOB" value=""  class="input-block-level" placeholder="DOB" />
<input type="text" name="Address" id="Address" value=""  class="input-block-level" placeholder="Address" />
<input type="text" name="City" id="City" value=""  class="input-block-level" placeholder="City" />
<input type="text" name="State" id="State" value=""  class="input-block-level" placeholder="State" />
<input type="text" name="Zip" id="Zip" value=""  class="input-block-level" placeholder="Zip" />
<input type="text" name="Email" id="Email" value=""  class="input-block-level" placeholder="Email" />
<input type="text" name="Cell Number" id="Cell Number" value=""  class="input-block-level" placeholder="Cell Number" />
    
<div class="actions">
<input type="submit" value="Register" name="submit" id="submitButton" class="btn btn-info" title="Register" />
</div>

</div>

</form>
</div>
              
    
              <!--#include file="layout/bottom-index.asp" -->
<script src="../../scripts/jquery.min.js" type="text/javascript"></script> 
<script src="../../scripts/bootstrap/js/bootstrap.min.js" type="text/javascript"></script>
<script src="../../scripts/default.js" type="text/javascript"></script>


<script src="../../scripts/carousel/jquery.carouFredSel-6.2.0-packed.js" type="text/javascript"></script><script type="text/javascript">$('#list_photos').carouFredSel({ responsive: true, width: '100%', scroll: 2, items: {width: 320,visible: {min: 2, max: 6}} });</script><script src="../../scripts/camera/scripts/camera.min.js" type="text/javascript"></script>
<script src="../../scripts/easing/jquery.easing.1.3.js" type="text/javascript"></script>
<script type="text/javascript">function startCamera() {$('#camera_wrap').camera({ fx: 'scrollLeft', time: 2000, loader: 'none', playPause: false, navigation: true, height: '50%', pagination: true });}$(function(){startCamera()});</script>


</body>
</html>
