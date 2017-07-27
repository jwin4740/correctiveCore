<%
'## included on every page.  Should be the common includer
%>

<!--#include file="adovbs.inc" -->
<!--#include file="dbconn.asp" -->
<!--#include file="variables.asp" -->
<!--#include file="functions.asp" -->

<%

If Session("setIsActive") = "" Then Call LoadHDSiteSettings()

%>