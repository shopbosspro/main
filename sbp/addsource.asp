<!-- #include file=aspscripts/auth.asp --> 
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style1 {
	color: #FFFFFF;
}
.style2 {
	background-image:url('newimages/wipheader.jpg')
}
.style3 {
	text-align: right;
}

-->
</style>
<script type="text/javascript">
function checkForm(){
	d = document.theform
	if(d.source.value.length == 0){
		alert("Source is required")
		return false;
	}


	document.theform.submit()
}

</script>

</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="style2" style="height: 51">
                    <tr>
                      <td height="11">
                       <p align="center" style="height: 51px">
						<br><span class="style1"><strong>Enter the new Source Here</strong></span></td>
                    </tr>
                  </table>
                </div>
               <form method="POST" action="addsourcenow.asp" name="theform">
                <div align="center">
                 <center>
                 <table border="0" width="50%">
                  <tr>
                   <td width="29%" class="style3" style="width: 50%"><b>Source:&nbsp; </b></td>
                   <td width="71%" style="width: 50%"><input type="text" name="source" size="20"></td>
                  </tr>
                 </table>
                 </center>
                </div>
                <p align="center"><input onmouseover=this.style.cursor='hand' type="button" onclick="checkForm()" value="Add Source" name="B1">
				<input onclick="location.href='sources.asp'" name="Button1" type="button" value="Cancel"></p>
               </form>
                &nbsp;</td>
            </tr>
            <tr>
              <td valign="top" width="100%"></td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
<p>&nbsp;</p>
<script language="javascript">document.theform.source.focus()</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
