<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->


<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" src="javascripts/killenter.js"></script>
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style1 {
	text-align: center;
}

.style2 {
	border: 1px solid #000000;
}
.style3 {
	font-weight: bold;
	border: 1px solid #000000;
}
.style4 {
	text-align: center;
	color: #FFFFFF;
	background-image: url('newimages/wipheader.jpg');
}

-->
</style>
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
              <td width="110" valign="top" align="left">
                &nbsp;</td>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style4"><strong>Enter your supplier data here</strong></td>
                    </tr>
                    </table>
                </div>
               <form name=theform action=addsuppliernow.asp method=post>
               <table border="0" style="width: 100%" cellpadding="3" cellspacing="0">
                <tr>
                 <td align="right" style="height: 26px; width: 42%;" class="style3">Name</td>
                 <td width="77%" style="height: 26px" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierName" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Address</td>
                 <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierAddress" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">City</td>
                 <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierCity" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">State</td>
                 <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierState" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Zip</td>
                 <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierZip" size="11" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Phone</td>
                 <td width="77%" class="style3"><font size="3">(<input onkeyup="if(document.theform.sac.value.length == 3){document.theform.spre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF">)
                  <input type="text" onkeyup="if(document.theform.spre.value.length == 3){document.theform.sl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="spre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF">-<input type="text" onkeyup="if(document.theform.sl4.value.length == 4){document.theform.sfac.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="sl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></font></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Fax</td>
                 <td width="77%" class="style3"><font size="3">(<input onkeyup="if(document.theform.sfac.value.length == 3){document.theform.sfpre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF">)
                  <input onkeyup="if(document.theform.sfpre.value.length == 3){document.theform.sfl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfpre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF">-<input onkeyup="if(document.theform.sfl4.value.length == 4){document.theform.SupplierContact.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></font></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Contact</td>
                 <td width="77%" class="style2"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="SupplierContact" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: #FFFFFF"></td>
                </tr>
                <tr>
                 <td align="right" class="style3" style="width: 42%">Active</td>
                 <td width="77%" class="style2"><select  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' size="1" name="Active">
                   <option value="YES" selected>YES</option>
                   <option value="NO">NO</option>
                  </select></td>
                </tr>
                <tr>
                 <td width="100%" colspan="2" class="style1"><input onmouseover=this.style.cursor='hand' onclick=document.theform.submit();window.opener.document.theform.Supplier.options[0].value=document.theform.SupplierName.value;window.opener.document.theform.Supplier.options[0].text=document.theform.SupplierName.value;self.close() type="button" value="Add Supplier" name="B1"></td>
                </tr>
               </table>
               <input type="hidden" name="supid" value="<%=request("supid")%>">
               </form>
              </td>
            </tr>
            <tr>
              <td width="110" valign="top" align="left" height="1">&nbsp;</td>
              <td valign="top" width="100%">&nbsp;</td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
<p>&nbsp;</p>
<script language="Javascript">
document.theform.SupplierName.focus()
</script>

</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
