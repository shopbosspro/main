<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->
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
	font-weight: bold;
	font-family: Arial, Helvetica, sans-serif;
	background-image: url('newimages/wipheader.jpg');
}
.style2 {
	border-width: 1px;
	font-weight: bold;
	border-style: solid;
}

.style3 {
	font-weight: bold;
	border-style: solid;
	background-image: url('newimages/pageheader.jpg');
}

-->
</style>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td style="height:50;" align="center" height="25" class="style1">
      <font size="3">Edit Vehicle Issues</font></td>
    </tr>
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td width="110" valign="top" align="left">
                <br>
                <br>
              </td>
              <td valign="top" width="100%"><p>&nbsp;<p>&nbsp;<div align="center">
                <center>
                <form method="POST" action="addcomplaintafteraction.asp">
                <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="0%" id="AutoNumber1">
                  <tr>
                    <td style="height:30;" width="100%" colspan="2" class="style3">
                    <p align="center"><font color="#FFFFFF">Edit Customer Concern 
                    Info Here</font></td>
                  </tr>
                  <tr>
                    <td width="37%" align="right" valign="top" class="style2">
                    <font size="2">Vehicle Issue:&nbsp; </font></td>
                    <td width="63%">
                    <textarea rows="5" name="complaint" cols="52" style="font-variant: small-caps; font-family: Arial; font-size: 12pt"><%=comp%></textarea></td>
                  </tr>
                  <tr>
                    <td width="37%" align="right" valign="top" class="style2">
                    <font size="2">Tech Story:&nbsp; </font></td>
                    <td width="63%">
                    <textarea rows="5" name="techreport" cols="52" style="font-variant: small-caps; font-family: Arial; font-size: 12pt"><%=tr%></textarea></td>
                  </tr>
                  <tr>
                    <td width="100%" colspan="2">
                    <p align="center">
                    <input onmouseover=this.style.cursor='hand' type="submit" value="Add Issue" name="B1"></td>
                  </tr>
                </table>
                <input type=hidden name=roid value="<%=request("roid")%>">
                </form>
                </center>
              </div>
              </td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
<p>&nbsp;</p>



</html>
<%
'Copyright 2011 - Boss Software Inc.
%>