<!-- #include file=aspscripts/conn.asp --> 
<!-- #include file=aspscripts/auth.asp --> 

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
if len(request("vin")) > 0 then
	stmt = "delete from vinscan where shopid = '" & shopid & "' and vin = '" & request("vin") & "'"
	'response.write stmt
	con.execute(stmt)
end if

'get userdefined fields from company
ustmt = "select CustomUserField1, CustomUserField2, CustomUserField3, usezipexplode from company where shopid = '" & request.cookies("shopid") & "'"
set urs = con.execute(ustmt)
urs.movefirst
cu1 = urs("CustomUserField1")
cu2 = urs("CustomUserField2")
cu3 = urs("CustomUserField3")
uze = lcase(urs("usezipexplode"))
set urs = nothing

%>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<script type="text/javascript" src="jquery/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="javascripts/bootstrap.min.js"></script>
<link rel="stylesheet" href="css/bootstrap.min.css">
<script language="javascript" type="text/javascript">

//*** launches the ajax connection

function getZip(z){
	<%
	if uze = "yes" then
	%>
	$.ajax({
		data: "zip="+z,
		url: "addcustzipdata.asp",
		success: function(r){
			if (r != "none*none"){
				car = r.split("*")
				city = car[0]
				st = car[1]
				$('#City').val(city)
				$('#State').val(st)
				$('#EMail').focus()
			}else{
				alert("No city or state found.  Possible bad zip code")
			}
		}
	});
	<%
	end if
	%>
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
		try{
			farray = document.theform.dl.value.split("^")
			strone = farray[0]
			document.theform.State.value = strone.substring(1,3)
			document.theform.City.value = strone.substring(3)
			sarray = farray[1].split("$")
			document.theform.LastName.value=sarray[0]
			document.theform.FirstName.value=sarray[1]
			document.theform.Address.value=farray[2]
			tarray = farray[3].split("!!")
			document.theform.Zip.value=Left(tarray[1],5)
			document.theform.HomePhone.value = ""
			document.theform.EMail.focus()
			document.theform.EMail.value = ""
			//document.theform.dl.style.display = "none"
		}catch(e){
			document.theform.dl.value = ""
			document.theform.dl.focus()
		}
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
	/*if(d.Address.value.length == 0){
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
	}*/


	document.theform.submit()
}

function checkDup(){

	xmlHttp=new XMLHttpRequest()
	trand = parseInt(Math.random()*99999999);  // cache buster
	fn = encodeURIComponent(document.theform.LastName.value)
	ln = encodeURIComponent(document.theform.FirstName.value)
	a = encodeURIComponent(document.theform.Address.value)
	url = "addcustcheckdup.asp?fn="+fn+"&ln="+ln+"&a="+a+"&shopid=<%=request.cookies("shopid")%>"
	url = url + "&myrand=" + trand
	console.log(url)
	xmlHttp.onreadystatechange=function(){
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
			r = xmlHttp.responseText
			console.log(r)
			if (r == "dup"){
				console.log("dup")
				c = confirm("This is a duplicate customer.  Click OK to be taken to the Customer Search screen or Cancel to enter the duplicate customer")
				if(c){
					top.location.href="wip.asp?vin=<%=request("vin")%>&ec=y&sp=<%=request("sp")%>&ln="+fn+"&fn="+ln
				}
			}
		}
	} 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function reswipe(){
	document.theform.LastName.value = ""
	document.theform.FirstName.value = ""
	document.theform.Address.value = ""
	document.theform.City.value = ""
	document.theform.State.value = ""
	document.theform.Zip.value = ""
	document.theform.dl.value = ""
	document.theform.dl.style.display = ""
	document.theform.EMail.value = ""
	document.theform.dl.focus()
}
<%if request("dl") = "y" then %>
x = setInterval(function(){
	d = document.theform.dl.value
	if (d.substring(d.length-1) == "?" && document.theform.LastName.value.length == 0){
		if(document.theform.EMail.value.indexOf("!") >= 0){
			document.theform.EMail.value = ""
		}
		decodeDL()
	}
},500)
<%end if%>
</script>

<style type="text/css">
<!--

p, td, th, li { font-size: 12pt; font-family: Verdana, Arial, Helvetica }
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

}
.auto-style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: small;
}
.auto-style2 {
	font-size: small;
}
.auto-style3 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
}
.auto-style4 {
	font-size: medium;
}
.auto-style5 {
	font-size: small;
	color: #FF0000;
}

.sbp-form-control{
	text-transform:uppercase;
	height:28px;padding:6px 12px;font-size:14px;color:#555;background-color:#fff;background-image:none;border:1px solid #ccc;border-radius:4px;
}

.form-control-phone{

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
<input type="hidden" name="vin" value="<%=request("vin")%>">

 <table border="0" cellpadding="0" cellspacing="0" style="width:100%;margin:auto" height="100%">
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
            <p align="center"><font color="#FFFFFF"><span class="auto-style4">Please enter all customer
            information, then click on the Add Customer button.</span><span class="auto-style2"><br>
			Note:&nbsp; We now check for duplicate customers.&nbsp; If you enter 
			a customer that has the same first name, last name and street 
			address as one already in your system, you will be redirected to the 
			Customer Search screen.</span></font><font size="2"><br>
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
          <input onblur="decodeDL()" name="dl" type="password" style="width:90%;background-color:yellow">
          <input name="rsbutton" type="button" value="Re-Swipe" onclick="reswipe()" >
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
        <table style="width: 90%;" cellspacing="0" cellpadding="2" class="table table-condensed">
         <tr>
          <td width="25%" align="right" class="auto-style1">
           <p align="right" class="auto-style1"><font size="3" color="#FF0000"><b>*</b></font>Last Name/Company Name:</p>
          </td>
          <td width="25%" align="left"><font size="1">
			<input type="text" onfocus="this.select()" value="<%if request("aid") > 0 then response.write ln else response.write request("LastName") end if%>" name="LastName" tabindex="1" class="sbp-form-control" size="24"></td>
          <td width="25%" align="right" class="auto-style1">
			Home
           Phone:</td>
          <td width="25%" align="left"><font size="1"><b>(<input onkeyup="if(document.theform.HomePhone.value.length == 3){document.theform.HomePhone1.focus()};" onblur="document.theform.HomePhone.style.backgroundColor='white'" onfocus="document.theform.HomePhone.style.backgroundColor='#FFFF99'" type="text" name="HomePhone" size="3" class="sbp-form-control form-control-phone"  tabindex="8" value="<%if request("aid") > 0 then response.write pn1 else response.write request("HomePhone") end if%>">)
           <input onkeyup="if(document.theform.HomePhone1.value.length == 3){document.theform.HomePhone2.focus()};" onblur="document.theform.HomePhone1.style.backgroundColor='white'" onfocus="document.theform.HomePhone1.style.backgroundColor='#FFFF99'" type="text" name="HomePhone1" size="3" class="sbp-form-control form-control-phone" tabindex="9" value="<%if request("aid") > 0 then response.write pn2 else response.write request("HomePhone1") end if%>">-<input onkeyup="if(document.theform.HomePhone2.value.length == 4){document.theform.WorkPhone.focus()};" onblur="document.theform.HomePhone2.style.backgroundColor='white'" onfocus="document.theform.HomePhone2.style.backgroundColor='#FFFF99'" type="text" name="HomePhone2" size="3" class="sbp-form-control form-control-phone"  tabindex="10" value="<%if request("aid") > 0 then response.write pn3 else response.write request("HomePhone2") end if%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style1">First
            Name:</td>
           <td width="25%" align="left"><font size="1"><input onkeyup="" onblur="document.theform.FirstName.style.backgroundColor='white'" onfocus="document.theform.FirstName.style.backgroundColor='#FFFF99'" type="text" name="FirstName" size="24" class="sbp-form-control" tabindex="2" value="<%if request("aid") > 0 then response.write fn else response.write request("FirstName") end if%>"></font></td>
           <td width="25%" align="right" class="auto-style1">Work
            Phone:</td>
           <td width="25%" align="left"><font size="1"><b>(<input onkeyup="if(document.theform.WorkPhone.value.length == 3){document.theform.WorkPhone1.focus()};" onblur="document.theform.WorkPhone.style.backgroundColor='white'" onfocus="document.theform.WorkPhone.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone" size="3" class="sbp-form-control form-control-phone" tabindex="11" value="<%=request("WorkPhone")%>">)
            <input onkeyup="if(document.theform.WorkPhone1.value.length == 3){document.theform.WorkPhone2.focus()};" onblur="document.theform.WorkPhone1.style.backgroundColor='white'" onfocus="document.theform.WorkPhone1.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone1" size="3" class="sbp-form-control form-control-phone" tabindex="12" value="<%=request("WorkPhone1")%>">-<input onkeyup="if(document.theform.WorkPhone2.value.length == 4){document.theform.CellPhone.focus()};" onblur="document.theform.WorkPhone2.style.backgroundColor='white'" onfocus="document.theform.WorkPhone2.style.backgroundColor='#FFFF99'" type="text" name="WorkPhone2" size="3" class="sbp-form-control form-control-phone" tabindex="13" value="<%=request("WorkPhone2")%>">&nbsp; 
		   Extension
		   <input type="text" name="WorkPhone3" size="3" class="sbp-form-control form-control-phone" tabindex="13" value="<%=request("WorkPhone3")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style1"><font size="3" color="#FF0000"><b>*</b></font>Address:</td>
           <td width="25%" align="left"><font size="1"><input onkeyup="" onblur="document.theform.Address.style.backgroundColor='white';" onfocus="document.theform.Address.style.backgroundColor='#FFFF99'" type="text" name="Address" size="24" class="sbp-form-control"  tabindex="3" value="<%=request("Address")%>"></font></td>
           <td width="25%" align="right" class="auto-style1">Cell
            Phone:</td>
           <td width="25%" align="left">
                     <font size="1"><b>(<input onkeyup="if(document.theform.CellPhone.value.length == 3){document.theform.CellPhone1.focus()};" onblur="document.theform.CellPhone.style.backgroundColor='white'" onfocus="document.theform.CellPhone.style.backgroundColor='#FFFF99'" type="text" name="CellPhone" size="3" class="sbp-form-control form-control-phone" tabindex="14" value="<%=request("CellPhone")%>">)
            <input onkeyup="if(document.theform.CellPhone1.value.length == 3){document.theform.CellPhone2.focus()};" onblur="document.theform.CellPhone1.style.backgroundColor='white'" onfocus="document.theform.CellPhone1.style.backgroundColor='#FFFF99'" type="text" name="CellPhone1" size="3" class="sbp-form-control form-control-phone" tabindex="15" value="<%=request("CellPhone1")%>">-<input onkeyup="if(document.theform.CellPhone2.value.length == 4){document.theform.cellprovider.focus()};" onblur="document.theform.CellPhone2.style.backgroundColor='white'" onfocus="document.theform.CellPhone2.style.backgroundColor='#FFFF99'" type="text" name="CellPhone2" size="3" class="sbp-form-control form-control-phone" tabindex="16" value="<%=request("CellPhone2")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style3">
			<strong><span class="auto-style2">*Zip</span></strong><span class="auto-style2">:<br></span>
			<span class="auto-style5"><strong>(If your computer freezes on this 
			box, <a onclick="showChromeError()" href="#" data-toggle="modal" data-target="#myModal">click here</a>)</strong></span></td>
           <td width="25%" align="left"><font size="3" color="#FF0000"><b><font size="1">
			<input onkeyup="" onblur="this.style.backgroundColor='white';getZip(this.value);" onfocus="document.theform.Zip.style.backgroundColor='#FFFF99'" type="text" name="Zip" size="14" class="sbp-form-control" tabindex="4" value="<%=request("Zip")%>"></font></b></font></td>
           <td width="25%" align="right" class="auto-style1">
			<strong>Spouse Name</strong></td>
           <td width="25%" align="left">
                     <font size="1">
			<input type="text" onfocus="this.select()" value="" name="spousename" tabindex="18" class="form-control" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style1">
			<font size="3" color="#FF0000"><b>*</b></font>City:</td>
           <td width="25%" align="left"><font size="3" color="#FF0000"><b><font size="1">
			<input onkeyup="" onblur="document.theform.City.style.backgroundColor='white'" onfocus="document.theform.City.style.backgroundColor='#FFFF99'" type="text" name="City" id="City" size="20" class="sbp-form-control"  tabindex="5" value="<%=request("City")%>"></font></b></font></td>
           <td width="25%" align="right" class="auto-style1">
			<strong>Spouse Work Phone</strong></td>
           <td width="25%" align="left">
                     <font size="1">
			<input type="text" onfocus="this.select()" value="" name="spousework" tabindex="19" class="form-control" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style1" rowspan="2">
			<font size="3" color="#FF0000"><b>
			*</b></font>State:</td>
           <td width="25%" align="left" rowspan="2"><font size="3" color="#FF0000"><b>
			<font size="1">
			<input onkeyup="" onblur="document.theform.State.style.backgroundColor='white'" onfocus="document.theform.State.style.backgroundColor='#FFFF99'" type="text" name="State" id="State" size="7" class="sbp-form-control" tabindex="6" value="<%=request("State")%>"></font></b></font></td>
           <td width="25%" align="right" class="auto-style3">
			<strong>Spouse Cell Phone</strong></td>
           <td width="25%" align="left"><font size="1">
			<input type="text" onfocus="this.select()" value="" name="spousecell" tabindex="20" class="form-control" size="24"></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style3">
			<strong>Fax:</strong></td>
           <td width="25%" align="left"><font size="1"><b>(<input onkeyup="if(document.theform.Fax.value.length == 3){document.theform.Fax1.focus()};" onblur="document.theform.Fax.style.backgroundColor='white'" onfocus="document.theform.Fax.style.backgroundColor='#FFFF99'" type="text" name="Fax" size="3" class="sbp-form-control form-control-phone" tabindex="21" value="<%=request("Fax")%>">)
            <input onkeyup="if(document.theform.Fax1.value.length == 3){document.theform.Fax2.focus()};" onblur="document.theform.Fax1.style.backgroundColor='white'" onfocus="document.theform.Fax1.style.backgroundColor='#FFFF99'" type="text" name="Fax1" size="3" class="sbp-form-control form-control-phone" tabindex="22" value="<%=request("Fax1")%>">-<input onkeyup="if(document.theform.Fax2.value.length == 4){document.theform.EMail.focus()};" onblur="document.theform.Fax2.style.backgroundColor='white'" onfocus="document.theform.Fax2.style.backgroundColor='#FFFF99'" type="text" name="Fax2" size="3" class="sbp-form-control form-control-phone" tabindex="23" value="<%=request("Fax2")%>"></b></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="auto-style1">
			<strong>E-Mail
            Address:</strong></td>
           <td width="25%" align="left"><font size="1">
		   <input onkeyup="" onblur="document.theform.EMail.style.backgroundColor='white'" onfocus="document.theform.EMail.style.backgroundColor='#FFFF99'" type="text" name="EMail" id="EMail" size="15" class="sbp-form-control" tabindex="7" value="<%=request("EMail")%>"></font></td>
           <td width="25%" align="right" class="auto-style3">
			<strong>Business Contact:</strong></td>
           <td width="25%" align="left"><font size="1">
			<input onkeyup="" onblur="document.theform.contact.style.backgroundColor='white'" onfocus="document.theform.contact.style.backgroundColor='#FFFF99'" type="text" name="contact" size="15" class="sbp-form-control" tabindex="24" value='<%=request("contact")%>'></font></td>
          </tr>
          <tr>
           <td width="25%" align="right" class="style4">
			&nbsp;</td>
           <td width="25%" align="left">&nbsp;</td>
           <td width="25%" align="right" class="auto-style3">
			<strong>Customer Pay Type:</strong> </td>
           <td width="25%" align="left">
		   <select name="customerpaytype" class="form-control"  tabindex="25">
		   <option value="Cash">Cash</option>
		   <option value="Net 10">Net 10</option>
		   <option value="Net 20">Net 20</option>
		   <option value="Net 30">Net 30</option>
		   </select></td>
          </tr>
          <%if len(cu1) > 0 then%>
          <tr>
           <td width="25%" align="right" class="style4">
			&nbsp;</td>
           <td width="25%" align="left">&nbsp;</td>
           <td width="25%" align="right" class="auto-style3">
			<strong>
			<%=cu1%>&nbsp;</strong></td>
           <td width="25%" align="left"><font size="1">
           
		   <input onkeyup="" onblur="document.theform.EMail.style.backgroundColor='white'" type="text" name="userdefined1" size="15" class="form-control" tabindex="26"></font>
		   
		   </td>
          </tr>
          <%end if%>
          <%if len(cu2) > 0 then%>
          <tr>
           <td width="25%" align="right" class="style4">
			&nbsp;</td>
           <td width="25%" align="left">&nbsp;</td>
           <td width="25%" align="right" class="auto-style3">
			<strong>
			<%=cu2%>&nbsp;</strong></td>
           <td width="25%" align="left"><font size="1">
           
		   <input onkeyup="" onblur="document.theform.EMail.style.backgroundColor='white'" type="text" name="userdefined2" size="15" class="form-control" tabindex="27"></font>
		   
		   </td>
          </tr>
          <%end if%>
          <%if len(cu3) > 0 then%>
          <tr>
           <td width="25%" align="right" class="style4">
			&nbsp;</td>
           <td width="25%" align="left">&nbsp;</td>
           <td width="25%" align="right" class="auto-style3">
			<strong>
			<%=cu3%>&nbsp;</strong></td>
           <td width="25%" align="left"><font size="1">
           
		   <input onkeyup="" onblur="document.theform.EMail.style.backgroundColor='white'" type="text" name="userdefined3" size="15" class="form-control" tabindex="28"></font>
		   
		   </td>
          </tr>
          <%end if%>
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
<div id="myModal" class="modal fade" role="dialog">
  <div class="modal-dialog">

    <!-- Modal content-->
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal">&times;</button>
        <h4 class="modal-title">If Chrome Freezes on Zip Code</h4>
      </div>
      <div class="modal-body">
        <p>You need to do 2 things.  First, you need to disable hardware acceleration.  Second, you need to disable Chrome's autofill function.  This should stop the freezing problem.
          Click the links below to see how.
          <br><br>
          <a href="https://www.lifewire.com/disable-form-autofill-in-chrome-446189" target="_blank">Disable Chrome Autofill</a>
          <br><br>
          <a href="http://ccm.net/faq/35743-google-chrome-how-to-disable-hardware-acceleration" target="_blank">Disable Hardware Acceleration</a>

          </p>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
      </div>
    </div>

  </div>
</div>
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
			response.write "getZip('" & rs("zip") & "');" & chr(10)
			//response.write "showCustomer('addcustzipdata.asp?zip=" & rs("zip") & "');" & chr(10)
		end if
			
	end if
	%>


</script>

<!-- #include file=helpfiles/customerentryhelp.asp -->
</html>
<%
'Copyright 2011 - Boss Software Inc.
%>