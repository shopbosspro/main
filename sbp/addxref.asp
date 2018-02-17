<!-- #include file=aspscripts/auth.asp --> 
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--
.menubutton{
	width:120px;
	height:66px;
	font-family:Arial, Helvetica, sans-serif;
	font-size:small;
	font-weight:bold;
	color:maroon;
	cursor:hand;
}
.menubutton:hover{
	color:blue;
}

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style1 {
	border: 0 solid #000000;
}
.style2 {
	font-weight: bold;
	border: 0 solid #000000;
}
.style3 {
	background-image: url('newimages/wipheader.jpg');
}
.style4 {
	color: #FFFFFF;
}

-->
</style>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style3">
                       <p align="center" class="style4"><strong>Enter the X Reference Number for this
                       part</strong></td>
                    </tr>
                  </table>
                </div>
               <div align="center">
                <center>
                <form name=theform action=addxrefnow.asp method=post>
                <table border="0" width="70%" cellpadding="3" cellspacing="0">
                 <tr>
                  <td width="43%" class="style2">Part Number</td>
                  <td width="57%" class="style1"><%=request("PartNumber")%></td>
                  <input type=hidden name=PartNumber value="<%=request("PartNumber")%>">
                 </tr>
                 <tr>
                  <td width="43%" class="style2">X Reference Part Number</td>
                  <td width="57%" class="style1"><input type="text" name="XREF" size="20"></td>
                 </tr>
                 <tr>
                  <td width="100%" bgcolor="#FFFFFF" colspan="2">
                   <p align="center"><input value="Add X-Reference" type="submit" name="Abutton1">&nbsp;<input onclick="history.go(-1)" name="Button1" type="button" value="Cancel"></p>
                  </td>
                 </tr>
                </table>
                </form>
                </center>
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
  </center>
</div>
<p>&nbsp;</p>



</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
