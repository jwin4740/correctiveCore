

<!DOCTYPE HTML>
<html>
<head>
    <meta charset="utf-8">
    <title>user login</title>
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
<br /><br/>
<div class="row-fluid">
<form action="jwUserLoginProcess.asp" method="post">
  
<div class="span6">
<input type="text" name="jwUsername" id="User Name" value=""  class="input-block-level" placeholder="User Name" />
<input type="password" name="jwPassword" id="Password" value=""  class="input-block-level" placeholder="Password" />

 <%If Session("jwUserLoginErrorMsg") <> "" Then %>
        <h5 style='margin-top: 0px; color: red; font-weight: bolder;' id="errorAlert"><%=Session("jwUserLoginErrorMsg")%></h5> 
    <%
        Session("jwUserLoginErrorMsg") = ""
    End If 
    %>
<div class="actions">
<input type="submit" value="Login" name="jwSubmit" id="submitButton" class="btn btn-info" title="Login" />
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
