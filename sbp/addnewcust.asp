<!-- #include file=aspscripts/conn.asp --> 
<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=javascripts/autocomplete.js -->

<%
if request("aid") > 0 then
	stmt = "select * from schedule where shopid = '" & request.cookies("shopid") & "' and id = " & request("aid")
	set rs = con.execute(stmt)
	rs.movefirst
	ln = rs("lastname")
	fn = rs("firstname")
	pn1 = left(rs("phonenumber"),3)
	pn2 = mid(rs("phonenumber"),4,3)
	pn3 = right(rs("phonenumber"),4)
	vyear = rs("year")
	vmake = rs("make")
	vmodel = rs("model")
	reason = rs("reason")
	set rs = nothing
	stmt = "delete from schedule where id = " & request("aid")
	con.execute(stmt)
end if

shopid = request.cookies("shopid")
lstmt = "select vinlabel vl, yearlabel yl,makelabel ml,modellabel modl,colorlabel cl,enginelabel el,cylinderlabel cyll,translabel tl,drivelabel dl,licenselabel ll,statelabel sl," _
& "fleetlabel fl,currmileagelabel cml from" _
& " vehiclelabels where shopid = '" & request.cookies("shopid") & "'"
set lrs = con.execute(lstmt)
yl = lrs("yl")
ml = lrs("ml")
modl = lrs("modl")
cl = lrs("cl")
el = lrs("el")
cyll = lrs("cyll")
tl = lrs("tl")
dl = lrs("dl")
ll = lrs("ll")
sl = lrs("sl")
fl = lrs("fl")
cml = lrs("cml")
vl = lrs("vl")
set lrs = nothing


' get custom vehicle fields
stmt = "select * from company where shopid = '" & shopid & "'"
set rs = con.execute(stmt)
carfax = lcase(rs("carfaxlocation"))
if len(rs("vehiclefield1label")) > 0 then
	vf1 = ucase(rs("vehiclefield1label"))
else
	vf1 = ""
end if
if len(rs("vehiclefield2label")) > 0 then
	vf2 = ucase(rs("vehiclefield2label"))
else
	vf2 = ""
end if
if len(rs("vehiclefield3label")) > 0 then
	vf3 = ucase(rs("vehiclefield3label"))
else
	vf3 = ""
end if
if len(rs("vehiclefield4label")) > 0 then
	vf4 = ucase(rs("vehiclefield4label"))
else
	vf4 = ""
end if
if len(rs("vehiclefield5label")) > 0 then
	vf5 = ucase(rs("vehiclefield5label"))
else
	vf5 = ""
end if
if len(rs("vehiclefield6label")) > 0 then
	vf6 = ucase(rs("vehiclefield6label"))
else
	vf6 = ""
end if
if len(rs("vehiclefield7label")) > 0 then
	vf7 = ucase(rs("vehiclefield7label"))
else
	vf7 = ""
end if
if len(rs("vehiclefield8label")) > 0 then
	vf8 = ucase(rs("vehiclefield8label"))
else
	vf8 = ""
end if


%>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<script type="text/javascript" src="javascripts/ajaxvin.js"></script>
<script type="text/javascript" src="javascripts/killenter.js"></script>
<script language="javascript" type="text/javascript">
function Left(str, n){
	if (n <= 0)     // Invalid bound, return blank string
		return "";
	else if (n > String(str).length)   // Invalid bound, return
		return str;                // entire string
	else // Valid bound, return appropriate substring
		return String(str).substring(0,n);
}
function Mid(str, start, len){
	if (start < 0 || len < 0) return "";

	var iEnd, iLen = String(str).length;
	if (start + len > iLen)
		iEnd = iLen;
	else
		iEnd = start + len;
		return String(str).substring(start,iEnd);
}
function pdmv(e){
	tstr = document.theform.Vin.value
	tstr = tstr.substring(0,2)
	if(tstr == "%0"){
		document.theform.License.value = Mid(document.theform.Vin.value,20,document.theform.Vin.value.length-19)
		document.theform.Vin.value = Mid(document.theform.Vin.value,2,17)
	}
	decodeVin(e)
}
function decodeVin(e){
	if(e == "c"){
		document.theform.Vin.onblur = "javascript:document.theform.Vin.style.backgroundColor='white'"
	}
	if(document.theform.Vin.value.length == 17){
		document.theform.Cyl.focus()
		document.theform.year.value = ""
		document.theform.Make.value = ""
		document.theform.Model.value = ""
		document.theform.Engine.value = ""
		document.theform.Cyl.value = ""
		var url=document.theform.Vin.value
		showList(url,"vin")
		document.theform.Vin.style.backgroundColor="white"
		clearTimeout(x)
		//alert(document.theform.Vin.value)
	}
}

function hpbar(){
t = setTimeout("hidepbar()",2000)
}

function hidepbar(){
document.getElementById('pbar').style.visibility='hidden';
document.getElementById('pbar').style.display='none'

}

function testLen(){
	if(document.theform.Vin.value.length > 20 && document.theform.Vin.value.indexOf("%0") == 0){
		setTimeout("pdmv('c')",750)
	}else{
		x = setTimeout("testLen()",100)
	}
}


//*** launches the ajax connection
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
		console.log(rt)
		if(rt != "none*none"){
			document.theform.EMail.focus()
			tarray = splitIt(rt,"*")
			document.theform.City.value=tarray[0]
			document.getElementById("state").value=tarray[1]
			console.log(tarray[1]+":"+document.getElementById("state").value)
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
function Left(str, n){
	if (n <= 0)     // Invalid bound, return blank string
		return "";
	else if (n > String(str).length)   // Invalid bound, return
		return str;                // entire string
	else // Valid bound, return appropriate substring
		return String(str).substring(0,n);
}
function Mid(str, start, len){
	if (start < 0 || len < 0) return "";

	var iEnd, iLen = String(str).length;
	if (start + len > iLen)
		iEnd = iLen;
	else
		iEnd = start + len;
		return String(str).substring(start,iEnd);
}

function decodeDL(){
	if(document.theform.dl.value.indexOf("^") > 0){
		farray = document.theform.dl.value.split("^")
		document.theform.State.value=Left(farray[0],2)
		document.theform.City.value=Mid(farray[0],2,farray[0].length-2)
		sarray = farray[1].split("$")
		document.theform.LastName.value=sarray[0]
		document.theform.FirstName.value=sarray[1]
		document.theform.Address.value=farray[2]
		tarray = farray[3].split("!!")
		document.theform.Zip.value=Left(tarray[1],5)
		document.theform.HomePhone.focus()
		document.theform.dl.style.display="none"
	}else{
		document.theform.dl.value = ""
	}
}
<%if request("dl") = "y" then %>
function testLen(){
	//alert(document.theform.dl.value)
	if(document.theform.dl.value.indexOf("^") > 0){
		decodeDL()
	}else{
		if(document.theform.dl.value.indexOf("%E") > 0){
			setTimeout("testLen()",1000)
			document.theform.dl.value = ""
		}
		setTimeout("testLen()",1000)
		document.theform.dl.value = ""
	}
}
<%end if%>

function checkForm(){
	d = document.theform
	if(d.LastName.value.length == 0){
		alert("Last Name is required")
		return false;
	}
	//if(d.FirstName.value.length == 0){
		//alert("First Name is required")
		//return false;
	//}
	if(d.Address.value.length == 0){
		alert("Street Address is required")
		return false;
	}
	if(d.City.value.length == 0){
		alert("City is required")
		return false;
	}
	if(d.State.value.length == 0){
		alert("State is required")
		return false;
	}
	if(d.Zip.value.length == 0){
		alert("Zip is required")
		return false;
	}


	document.theform.submit()
}
</script>

<style type="text/css">
<!--

p, td, th, li { font-size: 10pt; font-family: Verdana, Arial, Helvetica }
.style3 {
	font-weight: bold;
	background-image:url('newimages/wipheader.jpg')
}
.style4 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style5 {
	font-family: Arial, Helvetica, sans-serif;
}
.style6 {
	font-size: x-small;
}
.style7 {
	border: 1px solid black;
}
.style8 {
	border: 1px solid #203F80;
}
.auto-style1 {
	font-size: medium;
}
.auto-style2 {
	font-weight: bold;
}
.auto-style3 {
	text-align: center;
}
.auto-style4 {
	text-align: center;
	font-size: medium;
}
-->
</style>
</head>

<body  onload="
<%if request("dl") = "y" then %>
setTimeout('testLen()',100)
<%end if%>
" link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<!-- #include file=aspscripts/adovbs.asp -->
<br/><br/>
<form name="theform" action="addcustnow.asp" method="post">

<div align="center">
 <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
  <tr>
   <td width="100%" valign="top">
    <div align="left">
     <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
      <tr>
       <td valign="top" width="100%">
        <div align="center">
         <table style="width: 90%;" cellpadding="0" cellspacing="0" class="style7">
          <tr>
           <td height="11" class="style3">
            <p style="height:50px;" align="center"><font color="#FFFFFF">
			<span class="auto-style1">Please enter 
			all customer and vehicle information, then click on the Add button.</span></font><font size="2"><br>
            </font><font color="#FFFF00" size="1">* indicates required field
                        <%
                        if len(request("err")) > 0 then
                        	response.write "<br>You must enter information in all the required fields</font></b>"
                        end if
                       	%>
            </font></td>
          </tr>
          <%
          if request("dl") = "y" then
          %>
          
          <tr>
          <td width="100%">
          <input onblur="decodeDL()" name="dl" type="password" style="width:100%;background-color:yellow">
          </td>
          </tr>
          <%
          end if
          %>
         </table>
        </div>
	   
       
		<input type=hidden name=vmake value="<%=vmake%>">
		<input type=hidden name=vmodel value="<%=vmodel%>">
		<input type=hidden name=vyear value="<%=vyear%>">
		<input type=hidden name=reason value="<%=reason%>">
        <div align="center">
        <table style="width: 90%;" cellspacing="0" cellpadding="2" class="style8">
         <tr>
          <td width="25%" align="right" class="style4">
           <p align="right" class="style4"><font size="2">Last Name/Company 
		   Name:</font></p>
          </td>
          <td width="25%" align="left"><font size="1">
			<input type="text" onfocus="this.select()" value="<%if request("aid") > 0 then response.write ln else response.write request("LastName") end if%>" name="LastName" tabindex="1" style="overflow:hidden; font-size: 10pt; font-variant: small-caps; font-family: Arial; background-color: #FFFFFF; border-style: solid; border-color: gray; height: 20px;" size="24"><font size="3" color="#FF0000"><b>*</b></font></td>
          <td width="25%" align="right" class="style4">
			<font size="2">Home Phone:</font></td>
          <td width="25%" align="left"><font size="1"><b>(<input onkeyup="javascript:if(document.theform.HomePhone.value.length == 3){document.theform.HomePhone1.focus()};" onblur="javascript:document.theform.HomePhone.style.backgroundColor='white'" onfocus="javascript:document.theform.HomePhone.style.backgroundColor='#FFFF99'" type="text" name="HomePhone" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="8" value="<%if request("aid") > 0 then response.write pn1 else response.write request("HomePhone") end if%>">)
           <input onkeyup="javascript:if(document.theform.HomePhone1.value.length == 3){document.theform.HomePhone2.focus()};" onblur="javascript:document.theform.HomePhone1.style.backgroundColor='white'" onfocus="javascript:document.theform.HomePhone1.style.backgroundColor='#FFFF99'" type="text" name="HomePhone1" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="9" value="<%if request("aid") > 0 then response.write pn2 else response.write request("HomePhone1") end if%>">-<input onkeyup="javascript:if(document.theform.HomePhone2.value.length == 4){document.theform.WorkPhone.focus()};" onblur="javascript:document.theform.HomePhone2.style.backgroundColor='white'" onfocus="javascript:document.theform.HomePhone2.style.backgroundColor='#FFFF99'" type="text" name="HomePhone2" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray; height: 22px;" tabindex="10" value="<%if request("aid") > 0 then response.write pn3 else response.write request("HomePhone2") end if%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4"><font size="2">First 
		   Name:</font></td>
           <td width="25%" align="left"><font size="1"><input onkeyup="javascript:" onblur="javascript:document.theform.FirstName.style.backgroundColor='white'" onfocus="javascript:document.theform.FirstName.style.backgroundColor='#FFFF99'" type="text" name="FirstName" size="24" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; background-color: #FFFFFF; border-style: solid; border-color: gray" tabindex="2" value="<%if request("aid") > 0 then response.write fn else response.write request("FirstName") end if%>"></font></td>
           <td width="25%" align="right" class="style4"><font size="2">Work 
		   Phone:</font></td>
           <td width="25%" align="left"><font size="1"><b>(<input onkeyup="javascript:if(document.theform.WorkPhone.value.length == 3){document.theform.WorkPhone1.focus()};" onblur="javascript:document.theform.WorkPhone.style.backgroundColor='white'" onfocus="javascript:document.theform.WorkPhone.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="11" value="<%=request("WorkPhone")%>">)
            <input onkeyup="javascript:if(document.theform.WorkPhone1.value.length == 3){document.theform.WorkPhone2.focus()};" onblur="javascript:document.theform.WorkPhone1.style.backgroundColor='white'" onfocus="javascript:document.theform.WorkPhone1.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone1" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="12" value="<%=request("WorkPhone1")%>">-<input onkeyup="javascript:if(document.theform.WorkPhone2.value.length == 4){document.theform.CellPhone.focus()};" onblur="javascript:document.theform.WorkPhone2.style.backgroundColor='white'" onfocus="javascript:document.theform.WorkPhone2.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone2" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="13" value="<%=request("WorkPhone2")%>">&nbsp; 
		   Extension
		   <input type="text" name="WorkPhone3" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="13" value="<%=request("WorkPhone3")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4"><font size="2">Address:</font></td>
           <td width="25%" align="left"><font size="1"><input onkeyup="javascript:" onblur="javascript:document.theform.Address.style.backgroundColor='white'" onfocus="javascript:document.theform.Address.style.backgroundColor='#FFFF99'" type="text" name="Address" size="24" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; background-color: #FFFFFF; border-style: solid; border-color: gray" tabindex="3" value="<%=request("Address")%>"></font><font size="3" color="#FF0000"><b>*</b></font></td>
           <td width="25%" align="right" class="style4"><font size="2">Cell 
		   Phone:</font></td>
           <td width="25%" align="left">
                     <font size="1"><b>(<input onkeyup="javascript:if(document.theform.CellPhone.value.length == 3){document.theform.CellPhone1.focus()};" onblur="javascript:document.theform.CellPhone.style.backgroundColor='white'" onfocus="javascript:document.theform.CellPhone.style.backgroundColor='#FFFF99'" type="text" name="CellPhone" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="14" value="<%=request("CellPhone")%>">)
            <input onkeyup="javascript:if(document.theform.CellPhone1.value.length == 3){document.theform.CellPhone2.focus()};" onblur="javascript:document.theform.CellPhone1.style.backgroundColor='white'" onfocus="javascript:document.theform.CellPhone1.style.backgroundColor='#FFFF99'" type="text" name="CellPhone1" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="15" value="<%=request("CellPhone1")%>">-<input onkeyup="javascript:if(document.theform.CellPhone2.value.length == 4){document.theform.cellprovider.focus()};" onblur="javascript:document.theform.CellPhone2.style.backgroundColor='white'" onfocus="javascript:document.theform.CellPhone2.style.backgroundColor='#FFFF99'" type="text" name="CellPhone2" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="16" value="<%=request("CellPhone2")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style5">
			<strong>Zip</strong>:</td>
           <td width="25%" align="left"><font size="3" color="#FF0000"><b><font size="1">
			<input onblur="this.style.backgroundColor='white';showCustomer('addcustzipdata.asp?zip='+this.value);" type="text" name="Zip" size="14" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; background-color: #FFFFFF; border-style: solid; border-color: gray" tabindex="4" value="<%=request("Zip")%>"></font>*</b></font></td>
           <td width="25%" align="right" class="style4">
			<strong>Spouse Name</strong></td>
           <td width="25%" align="left">
                     <font size="1">
			<input type="text" onfocus="this.select()" value="<%if request("aid") > 0 then response.write ln else response.write request("LastName") end if%>" name="spousename" tabindex="18" style="overflow:hidden; font-size: 10pt; font-variant: small-caps; font-family: Arial; background-color: #FFFFFF; border-style: solid; border-color: gray; height: 20px;" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4">
			<font size="2">City:</font></td>
           <td width="25%" align="left"><font size="3" color="#FF0000"><b><font size="1">
			<input onkeyup="javascript:" onblur="javascript:document.theform.City.style.backgroundColor='white'" onfocus="javascript:document.theform.City.style.backgroundColor='#FFFF99'" type="text" name="City" size="20" style="font-size: 10pt; font-variant: small-caps; font-family: Arial; background-color: #FFFFFF; border-style: solid; border-color: gray" tabindex="5" value="<%=request("City")%>"></font>*</b></font></td>
           <td width="25%" align="right" class="style4">
			<strong>Spouse Work Phone</strong></td>
           <td width="25%" align="left">
                     <font size="1">
			<input type="text" onfocus="this.select()" value="<%if request("aid") > 0 then response.write ln else response.write request("LastName") end if%>" name="spousework" tabindex="19" style="overflow:hidden; font-size: 10pt; font-variant: small-caps; font-family: Arial; background-color: #FFFFFF; border-style: solid; border-color: gray; height: 20px;" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4" rowspan="2">
			<font size="2">State:</font></td>
           <td width="25%" align="left" rowspan="2"><font size="3" color="#FF0000"><b>
			<font size="1">
			<input id="state" type="text" name="State" size="7" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; background-color: #FFFFFF; border-style: solid; border-color: gray" tabindex="6" value="<%=request("State")%>"></font>*</b></font></td>
           <td width="25%" align="right" class="style5">
			<strong>Spouse Cell Phone</strong></td>
           <td width="25%" align="left"><font size="1">
			<input type="text" onfocus="this.select()" value="<%if request("aid") > 0 then response.write ln else response.write request("LastName") end if%>" name="spousecell" tabindex="20" style="overflow:hidden; font-size: 10pt; font-variant: small-caps; font-family: Arial; background-color: #FFFFFF; border-style: solid; border-color: gray; height: 20px;" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style5">
			<font size="2"><strong>Fax:</strong></font></td>
           <td width="25%" align="left"><font size="1"><b>(<input onkeyup="javascript:if(document.theform.Fax.value.length == 3){document.theform.Fax1.focus()};" onblur="javascript:document.theform.Fax.style.backgroundColor='white'" onfocus="javascript:document.theform.Fax.style.backgroundColor='#FFFF99'" type="text" name="Fax" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="21" value="<%=request("Fax")%>">)
            <input onkeyup="javascript:if(document.theform.Fax1.value.length == 3){document.theform.Fax2.focus()};" onblur="javascript:document.theform.Fax1.style.backgroundColor='white'" onfocus="javascript:document.theform.Fax1.style.backgroundColor='#FFFF99'" type="text" name="Fax1" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="22" value="<%=request("Fax1")%>">-<input onkeyup="javascript:if(document.theform.Fax2.value.length == 4){document.theform.EMail.focus()};" onblur="javascript:document.theform.Fax2.style.backgroundColor='white'" onfocus="javascript:document.theform.Fax2.style.backgroundColor='#FFFF99'" type="text" name="Fax2" size="3" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray" tabindex="23" value="<%=request("Fax2")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4">
			<font size="2"><strong>E-Mail Address:</strong></font></td>
           <td width="25%" align="left"><font size="1">
		   <input onblur="javascript:document.theform.EMail.style.backgroundColor='white'" onfocus="javascript:document.theform.EMail.style.backgroundColor='#FFFF99'" type="text" name="EMail" size="15" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray; width: 184px;" tabindex="7" value="<%=request("EMail")%>"></font></td>
           <td width="25%" align="right" class="style5">
			<strong>Business Contact:</strong></td>
           <td width="25%" align="left"><font size="1">
			<input onkeyup="javascript:" onblur="javascript:document.theform.contact.style.backgroundColor='white'" onfocus="javascript:document.theform.contact.style.backgroundColor='#FFFF99'" type="text" name="contact" size="15" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray; height: 20px;" tabindex="24" value='<%=request("contact")%>'></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4">
			&nbsp;</td>
           <td width="25%" align="left">&nbsp;</td>
           <td width="25%" align="right" class="style5">
			<strong>Customer Pay Type:</strong> </td>
           <td width="25%" align="left"><select name="customerpaytype">
		   <option value="Cash">Cash</option>
		   <option value="Net 10">Net 10</option>
		   <option value="Net 20">Net 20</option>
		   <option value="Net 30">Net 30</option>
		   </select></td>
          </tr>
          <tr>
           <td align="right" class="style4" colspan="4">
           <hr>
		   <div class="auto-style4">
			   Vehicle Information
</div>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
      <tr>
       <td valign="top" width="100%">
        <div align="center">
         <center>
         <input type="hidden" name="origshopid" value="<%=request("origshopid")%>">
         <input type="hidden" name="reason" value="<%=request("reason")%>">
          <table cellspacing="0" cellpadding="2" class="style1" style="width: 70%">
           <tr>

            <td class="auto-style2" style="width: 20%"><%=vl & vstar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="" onblur="javascript:document.theform.Vin.style.backgroundColor='white'" onfocus="javascript:document.theform.Vin.style.backgroundColor='#FFFF99'" type="text" name="Vin" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Vin")%>">
			<input type="button" tabindex="15" onclick="pdmv('c')" name="Abutton1" value="Click to Decode">&nbsp;<span class="style5"><strong>(1981 
			newer)</strong></span></td>
            <td  class="auto-style2" style="width: 60%;padding:10px;" rowspan="21" valign="top">
			<span class="auto-style3"><strong>*** NOTE ***</strong></span><strong><br>
			<br>Vehicle fields can now be customized <em>AND </em>marked as 
			required or not required.<br><br><br>To customize a label for a vehicle 
			field, go to Settings -&gt; Custom.&nbsp; There you can change the name 
			of any field and mark them as required or not required.<br><br><br>You can also add additional 
			custom vehicle fields to save information other than the standard 
			fields.</strong></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=yl & ystar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="javascript:" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="year" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Year")%>">
				<div style="padding-left:4px;padding-right:4px;padding-top:2px;height:22px;background-color:black;margin-top:-22px;width:250px;display:none" id="pbar">
					<font color="red"><b>Decoding VIN.  Please stand by...</b></font></div></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=ml & mstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="Make" onkeyup="javascript:" onblur="this.style.backgroundColor='white';" onfocus="javascript:loadMakes(this,mlist,document.theform.Model);this.style.backgroundColor='#FFFF99'" type="text" name="Make" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray; width: 157px;" value='<%=request("make")%>'></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=modl & modstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="Model" onkeyup="javascript:" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:loadMakes(this,document.theform.Make.value,document.theform.Engine);this.style.backgroundColor='#FFFF99'" type="text" name="Model" size="20" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("model")%>'>
			</td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=cyll & cylstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="cyl" onkeyup="javascript:" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:loadMakes(this,document.theform.Make.value,document.theform.Engine);this.style.backgroundColor='#FFFF99'" type="text" name="Cyl" size="20" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("cyl")%>'></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=el & estar%></td>
            <td  class="style4" style="width: 60%">
			<input onfocus="this.style.backgroundColor='#FFFF99'" onblur="this.style.backgroundColor='white'" onkeyup="javascript:if (this.value.length == 3){autoComplete(this.value, document.theform.Engine)};" type="text" name="Engine" size="14" value="<%=request("Engine")%>" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray">
             (ie. 2.4 Liter)</td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=cl & cstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="color" onkeyup="javascript:" onblur="javascript:document.theform.Cyl.style.backgroundColor='white'" onfocus="javascript:document.theform.Cyl.style.backgroundColor='#FFFF99';hpbar()" type="text" name="Color" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Cyl")%>"></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=tl & tstar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="javascript:" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="Trans" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Trans")%>"> 
			(AUTO, MAN, 5SP)</td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%; height: 24px;"><%=dl & dstar%></td>
            <td  class="style4" style="width: 60%; height: 24px;">
			<input onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="drivetype" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("fleetno")%>'> 
			(RWD, FWD, AWD)</td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=ll & lstar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="javascript:" onblur="javascript:document.theform.License.style.backgroundColor='white'" onfocus="javascript:document.theform.License.style.backgroundColor='#FFFF99'" type="text" name="License" size="14" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("License")%>"></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=sl & sstar%></td>
            <td  class="style4" style="width: 60%">
			<input onblur="this.style.backgroundColor='white'" onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" type="text" name="State" size="6" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray; background-color:white" value="<%=shopstate%>"></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=fl & fstar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="javascript:" 
			onblur="javascript:document.theform.State.style.backgroundColor='white'" 
			onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" 
			type="text" name="fleetno" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value='<%=request("fleetno")%>'></td>
           </tr>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=cml & cmstar%></td>
            <td  class="style4" style="width: 60%">
			<input onkeyup="javascript:" 
			onblur="javascript:document.theform.State.style.backgroundColor='white'" 
			onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" 
			type="text" name="currentmileage" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value='<%=request("currentmileage")%>'></td>
           </tr>
           <%	if len(vf1) > 0 then	%>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf1%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom1" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%
           end if
           if len(vf2) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf2%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom2" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%
           end if
           if len(vf3) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf3%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom3" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%
           end if
           if len(vf4) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf4%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom4" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%	
           end if	
           
           if len(vf5) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf5%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom5" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%	
           end if	
           
           if len(vf6) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf6%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom6" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%	
           end if	
           
           if len(vf7) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf7%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom7" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%	
           end if	
           
           if len(vf8) > 0 then
           %>
           <tr>
            <td class="auto-style2" style="width: 20%"><%=vf8%></td>
            <td  class="style4" style="width: 60%">
			<input 
			type="text" name="custom8" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value=''></td>
           </tr>
           <%	
           end if	
           
           %>

			<tr>
				<td colspan="3" class="auto-style1"><strong><br>Need more fields?  Go to Settings -&gt; Custom.  There you can add up to 8 additional custom fields</strong></td>
			</tr>



           <tr>
            <td  colspan="3" align="right" class="style2" style="width: 100%">
             <p align="center"><input type="button" value="Add Vehicle" onclick="subForm()" class="obutton" name="Abutton2">
			 <input type="button" value="Back to WIP" onclick="location.href='wip.asp'" class="obutton" name="Abutton3">
			 <input type="button" value="Create Part Sale" onclick="location.href='aspscripts/createps.asp?id=<%=request("custid")%>'" class="obutton" name="Abutton4"></td>
           </tr>
           </table>
			<input type="hidden" name="sp" value="<%=request("sp")%>">
          <input type="hidden" name="custid" value="<%=request("custid")%>">
         <input type="hidden" name="appt" value="<%=request("appt")%>">
         <input type="hidden" name="d" value="<%=request("d")%>">
         </form>
         </center>
        </div>
       </td>
      </tr>
      </table>

			</td>
          </tr>
          <tr>
           <td width="100%" align="right" colspan="4" style="height: 33px">
            <p align="center"><input style="cursor:pointer" onclick="checkForm()" type="button" value="Add Customer">
			<input type="button" value="Cancel" onclick="history.go(-1)" name="Abutton1"></td>
          </tr>
         </table>
          </div>
         <input type="hidden" name="vehnum" value><input type="hidden" name="sp" value="addcust.asp">
        <input type="hidden" name="appt" value="<%=request("appt")%>">
        <input type="hidden" name="d" value="<%=request("d")%>">

</td>
     </tr>
     </table>
   </div>
  </td>
 </tr>
 </table>
</div>
<p>&nbsp;</p>
</form>
<script language="javascript">
	<%if request("dl") = "y" then%>
	document.theform.dl.focus();
	<%else%>
	document.theform.LastName.focus();
	<%end if%>
</script>
<script type="text/javascript">
	<%
	set rs = nothing
	if len(request.cookies("quoteid")) > 0 then
		stmt = "select * from quotes where shopid = '" & request.cookies("shopid") & "' and id = " & request.cookies("quoteid")
		set rs = con.execute(stmt)
		'response.write stmt
		response.write "document.theform.LastName.value = '" & rs("customer") & "';" & chr(10)
		response.write "document.theform.Address.value = '" & rs("address") & "';" & chr(10)
		response.write "document.theform.Zip.value = '" & rs("zip") & "';" & chr(10)
		if len(rs("zip")) > 0 then
			response.write "showCustomer('addcustzipdata.asp?zip=" & rs("zip") & "');" & chr(10)
		end if
			
	end if
	%>


</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>