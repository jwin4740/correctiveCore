<script type="text/javascript" language="javascript" >
    function ShowProgress()
    {
      strAppVersion = navigator.appVersion;
      if (document.<%=hdThisFormName%>.<%=hdThisFileFieldName%>.value != "")
      {
        if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
        {
	    winstyle = "dialogWidth=385px; dialogHeight:140px; center:yes";
	    window.showModelessDialog('<%= barref %>&b=IE', null, winstyle);
        }
        else
        {
          window.open('<%= barref %>&b=NN','','width=375,height=115', true);
        }
      }
      return true;
    }
</script>