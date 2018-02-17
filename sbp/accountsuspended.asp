<%
response.redirect "renewaccount.asp?shopid=" & request("shopid")
%>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/main.css">
<title><%=request.cookies("shopname")%></title>

<style type="text/css">
.style1 {
	font-size: small;
}
.style2 {
	font-size: large;
}
.auto-style1 {
	font-size: large;
	color: #800000;
}
#renewbox:hover{
	background-color:#006600;
	color:white;
}
.auto-style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: medium;
}
.auto-style3 {
	font-size: medium;
}
.auto-style4 {
	color: #800000;
}
</style>

</head>

<body >


<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td align="center" height="25" class="sbheader">
        <img border="0" src="newimages/shopbosslogo.png">
      </td>
    </tr>
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                      <td height="11">
                        <p align="center">
                        <p align="center">
                        &nbsp;<p align="center" class="style2"><!--onclick="location.href='renewaccount.asp?shopid=<%=request("shopid")%>'" -->
                        &nbsp;<p id="renewbox" onmouseover="this.style.backgroundColor='#006600'" onmouseout="this.style.backgroundColor='#336699'" style="cursor:pointer;background-color:#336699;color:white;margin-left:30%;width:40%;border:5px black solid;border-radius:20px;padding:40px;-webkit-box-shadow: 10px 10px 5px 0px rgba(0,0,0,0.75);-moz-box-shadow: 10px 10px 5px 0px rgba(0,0,0,0.75);box-shadow: 10px 10px 5px 0px rgba(0,0,0,0.75);" align="center">
                        <span class="style2"><strong>We're sorry but we were unable to charge your 
						credit card for your subscription<br><br>Please contact 
						support at<br>858-435-0227 to update your credit card 
						information</strong></span><br><strong>
						<span style="text-transform:uppercase" class="auto-style1">
						<p align="center">
                        &nbsp;<p align="center">
                        &nbsp;</span><span style="text-transform:uppercase" class="auto-style4"><p align="center">
                        <span class="auto-style2">If you have any questions, please contact your account representative by
						<a href="mailto:support@shopbosspro.com"><strong>email</strong></a> 
						or 
						phone at </span></span><span style="text-transform:uppercase" class="auto-style1">
						<font color="#000000" face="Arial, Helvetica, sans-serif" class="auto-style3">
						(858) 435-0227<strong> </strong>or <a href="login1.asp">Click Here</a> 
						to re-try your logIn</font></td>
                    </tr>
                  </table>
                </div>
                </td>
            </tr>
            <tr>
              <td valign="top" width="100%"></td>
            </tr>
            </table>
        </div>
    </td>
    </tr>
  </table>
</div>
</form>
</html>
<%
'Copyright 2011 - Boss Software Inc.
%>