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
	text-align: center;
	color: #FFFFFF;
		background-image: url('newimages/wipheader.jpg');
}
.style2 {
	text-align: right;
}

-->
</style>
<script type="text/javascript">

function checkForm(){
	if(document.theform.JobDesc.value.length == 0){
		alert("Job Description is required")
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
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style1">
                       <strong>Enter the new Job Description Here</strong></td>
                    </tr>
                  </table>
                </div>
               <form method="POST" action="addjobdescnow.asp" name="theform">
                <div align="center">
                 <center>
                 <table border="0" width="50%">
                  <tr>
                   <td width="29%" class="style2" style="width: 50%"><b>Job Description: </b></td>
                   <td width="71%" style="width: 50%"><input type="text" name="JobDesc" size="20"></td>
                  </tr>
                 </table>
                 </center>
                </div>
                <p align="center"><input onmouseover=this.style.cursor='hand' type="button" onclick="checkForm()" value="Add Job Description" name="B1">
				<input onclick="location.href='jobdescs.asp'" name="Button1" type="button" value="Cancel"></p>
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
<script language="javascript">document.theform.JobDesc.focus()</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
