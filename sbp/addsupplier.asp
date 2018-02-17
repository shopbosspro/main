<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->


<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<script type="text/javascript" src="javascripts/killenter.js"></script>
<script type="text/javascript">

function showCustomer(str){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url=str
	trand = parseInt(Math.random()*99999999);  // cache buster
	url = url + "&myrand=" + trand
	//alert(url)
	xmlHttp.onreadystatechange=stateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}
	
	//*** splits the return string
function splitIt(splitItStr,splItter){
	myarray = splitItStr.split("*");
	return myarray;
}

var xmlHttp

//*** state changed, is the xml response text ready
function stateChanged(){ 
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
		rt = xmlHttp.responseText
		if(rt != "none*none"){
			document.theform.sac.focus()
			tarray = splitIt(rt,"*")
			document.theform.SupplierCity.value=tarray[0]
			document.theform.SupplierState.value=tarray[1]
		}else{
			alert("No city or state found.  Possible bad zip code")
		}
	} 
}

//*** creates the xml object
function GetXmlHttpObject(){ 
	var objXMLHttp=null
	if (window.XMLHttpRequest){	
		objXMLHttp=new XMLHttpRequest()
	}	
	else if (window.ActiveXObject)
	{
	objXMLHttp=new ActiveXObject("Microsoft.XMLHTTP")
	}
	return objXMLHttp
}

</script>
<style type="text/css">
<!--

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style1 {
	border: 1px solid #B6CDFB;
}
.style2 {
	font-weight: bold;
	border: 0 solid #000000;
}

.style3 {
	color: #FFFFFF;
}
.style4 {
	font-weight: bold;
	border: 0 solid #B6CDFB;
}

.style5 {
	background-image: url('newimages/wipheader.jpg');
}
.style7 {
	border: 0 solid #000000;
}

-->
</style>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td height="25" align="center" height="25" class="style5">
		<span class="style3"><strong><br>Add 
		Supplier Information below.</strong></span><br>
		<br>
		</td>
    </tr>
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
               <form name=theform action=addsuppliernow.asp method=post>
               <table cellspacing="0" cellpadding="3" style="width: 70%" class="style1" align="center">
                <tr>
                 <td width="23%" align="right" class="style4" colspan="2" style="width: 100%">
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Name</td>
                 <td width="77%" class="style7"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierName" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Address</td>
                 <td width="77%" class="style7"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierAddress" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Zip</td>
                 <td width="77%" class="style7"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur="this.style.backgroundColor='white';showCustomer('addcustzipdata.asp?zip='+this.value);" name="SupplierZip" size="11" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">City</td>
                 <td width="77%" class="style7"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierCity" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">State</td>
                 <td width="77%" class="style7"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierState" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Phone</td>
                 <td width="77%" class="style2"><font size="3">(<input onkeyup="if(document.theform.sac.value.length == 3){document.theform.spre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">)
                  <input type="text" onkeyup="if(document.theform.spre.value.length == 3){document.theform.sl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="spre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">-<input type="text" onkeyup="if(document.theform.sl4.value.length == 4){document.theform.sfac.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="sl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></font></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Fax</td>
                 <td width="77%" class="style2"><font size="3">(<input onkeyup="if(document.theform.sfac.value.length == 3){document.theform.sfpre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">)
                  <input onkeyup="if(document.theform.sfpre.value.length == 3){document.theform.sfl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfpre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">-<input onkeyup="if(document.theform.sfl4.value.length == 4){document.theform.SupplierContact.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></font></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Contact</td>
                 <td width="77%" class="style7"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="SupplierContact" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Active</td>
                 <td width="77%" class="style7"><select  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' size="1" name="Active">
                   <option value="YES" selected>YES</option>
                   <option value="NO">NO</option>
                  </select></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Account #</td>
                 <td width="77%" class="style7">
				 <input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="accountnumber" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Terms</td>
                 <td width="77%" class="style7">
				 <input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="terms" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Tax Exempt</td>
                 <td width="77%" class="style7">
				 <input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="taxexempt" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
                </tr>
                <tr>
                 <td width="23%" align="right" class="style2">Notes:</td>
                 <td width="77%" class="style7">
				 <textarea name="notes" style="width: 565px; height: 65px"></textarea></td>
                </tr>
                <tr>
                 <td width="100%" colspan="2" align="center"><input onmouseover=this.style.cursor='hand' onclick=document.theform.submit() type="button" value="Add Supplier" name="B1">
		<input type="button" value="Cancel" onclick="location.href='suppliers.asp'" name="Abutton1"></td>
                </tr>
               </table>
               <input type="hidden" name="supid" value="<%=request("supid")%>">
               	<input name="sp" type="hidden" value="suppliers.asp">
               </form>
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
<script language="Javascript">
document.theform.SupplierName.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
