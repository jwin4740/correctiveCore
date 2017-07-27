<!--#include file="hddriversqlinc.asp" -->

<%

'## if page found and no errors, display the page
If bContinue Then

    '## get first news record.
    Set rsNews = Server.CreateObject("ADODB.Recordset")
    rsNews.ActiveConnection = hdDSN
    rsNews.Source = "SELECT TOP 1 * FROM hdNews ORDER BY newsID DESC"
    rsNews.Open()
%>

<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">

<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title><%=webpgTitle%></title>

<meta name="description" content="<%=webpgMetaDescription%>" />

<meta name="keywords" content="<%=webpgMetaKeywords%>" />

<meta property="og:url" content="http://www.correctivecore.net/" />
<meta property="og:site_name" content="Corrective Core" />
<meta property="og:title" content="<%=webpgTitle%>" />
<meta property="og:description" content="<%=webpgMetaDescription%>" />

<link href='http://tychem.net/favicon.ico' rel='icon' type='image/x-icon'/>
<meta name="author" content="Corrective Core">
<meta name="robots" content="index, follow" />

<link rel="canonical" href="http://www.correctivecore.net/<%=webpgFileName%>" />

<meta property="og:locale" content="en_US" />
<meta property="og:type" content="website" />
<meta property="og:title" content="<%=webpgTitle%>" />
<meta property="og:description" content="<%=webpgMetaDescription%>" />
<meta property="og:url" content="http://www.correctivecore.net/<%=webpgFileName%>" />
<meta property="og:site_name" content="North Carolina" />

<script type="application/ld+json">
{
  "@context": "http://schema.org",
  "@type": "website",
  "name": "Corrective Core",
  "description": "<%=webpgMetaDescription%>",
  "url": "http://www.correctivecore.net/<%=webpgFileName%>",
  "contactPoint": [{
    "@type": "ContactPoint",
    "telephone": "+1-980-434-6770",
    "contactType": "Schedule Appointment"
  }]
}
</script>

<link href="scripts/bootstrap/css/bootstrap.min.css" rel="stylesheet">
<link href="scripts/bootstrap/css/bootstrap-responsive.min.css" rel="stylesheet">
<link href="scripts/icons/general/stylesheets/general_foundicons.css" media="screen" rel="stylesheet" type="text/css" />  
<link href="scripts/icons/social/stylesheets/social_foundicons.css" media="screen" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="scripts/fontawesome/css/font-awesome.min.css">
<link href="scripts/carousel/style.css" rel="stylesheet" type="text/css" />
<link href="scripts/camera/css/camera.css" rel="stylesheet" type="text/css" />

<link href="http://fonts.googleapis.com/css?family=Allura" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Aldrich" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Pacifico" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Palatino+Linotype" rel="stylesheet" type="text/css">
<link href="http://fonts.googleapis.com/css?family=Calligraffitti" rel="stylesheet" type="text/css">

<link href="styles/custom.css" rel="stylesheet" type="text/css" />

</head>
<body id="pageBody">

<!--#include file="layout/header.asp" -->
<!--#include file="layout/banner-thrive.asp" -->
<div id="footerInnerSeparator">

<div class="container">
<div class="divPanel page-content">
<div class="row-fluid">
<div class="span12" id="divMain">
<% Call PageContent %>
</div>
</div>
</div>
</div>


<div id="footerOuterSeparator">
<div id="decorative5">
<div class="container">
<div class="span12">
<h1 class="decorative5">Have questions?<br />
Call Or Text Today 980-434-6770</h1>
</div>
</div>
</div>
<!--#include file="layout/footer.asp" -->


<script src="scripts/jquery.min.js" type="text/javascript"></script> 
<script src="scripts/bootstrap/js/bootstrap.min.js" type="text/javascript"></script>
<script src="scripts/default.js" type="text/javascript"></script>

</body>
</html>

<%

    rsNews.Close
    Set rsNews = Nothing

End If  '## bContinue

'## error on db read, bail
If Not bContinue Then Response.Redirect("error.asp")

Function getRsNews(hdTotalRecords, theCategory)

    hdSQL = "SELECT TOP " & hdTotalRecords & " hdNews.newsIsFile, hdNews.newsID, " & _
        "hdNews.newsDetails, hdNews.newsShortDesc, hdNews.newsTitle FROM hdCategories INNER JOIN hdNews " & _
        "ON hdCategories.catID = hdNews.catID  WHERE hdCategories.cattypeID = " & hdNEWScat & _
        " AND hdCategories.catName = '" & theCategory & "' ORDER BY hdNews.newsID DESC"

    Set rsgetRsNews = Server.CreateObject("ADODB.Recordset")
    rsgetRsNews.ActiveConnection = hdDSN
    rsgetRsNews.Source = hdSQL
    rsgetRsNews.Open()
    If rsgetRsNews.EOF Then
        getRsNews = ""
    Else
        getRsNews = rsgetRsNews.GetRows
    End If
    
    rsgetRsNews.Close()
    Set rsgetRsNews = Nothing    
    
End Function

%>

