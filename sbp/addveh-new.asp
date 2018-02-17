<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=javascripts/autocomplete.js -->
<!-- #include file=javascripts/ajaxvin.js -->
<!-- #include file=aspscripts/auth.asp -->
<%
ua = request.servervariables("http_user_agent")
if instr(ua,"Android") > 0 then android = "yes"
if instr(ua,"iPod") > 0 then ipod = "yes"
if instr(ua,"iPad") > 0 then ipad = "yes"
if instr(ua,"iPhone") > 0 then iphone = "yes"
if len(request.querystring("origshopid")) > 0 then
	shopid = request.querystring("origshopid")
else
	shopid = request.cookies("shopid")
end if

if len(request("state")) > 0 then
	stopstate = request("state")
else
	shopstate = trim(request.cookies("shopstate"))
end if

if len(request("vin")) > 17 then
	tvin = request("vin")
	tvar = split(tvin,"|")
	nvin = left(replace(tvar(0)," ",""),17)
	custid = tvar(1)
	sp = tvar(2)
	
	response.redirect "addveh-new.asp?custid=" & custid & "&vin=" & nvin & "&sp=" & sp
end if


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
<script language="javascript" type="text/javascript" src="javascripts/actb.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/actbcommon.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/modelarray.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/modelarray.js"></script>
<script type="text/javascript" src="jquery/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="javascripts/bridgeit.js"></script>

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
</script>
<SCRIPT LANGUAGE="JavaScript">
//suppress error messages
//arrays of models
//array of makes
<%
av = "var mlist = new Array("
stmt = "select distinct make from model order by make"
set rs = con.execute(stmt)
rs.movefirst
do until rs.eof
	av = av & "'" & rs("make") & "',"
	rs.movenext
loop
set rs = nothing
av = left(av,len(av)-1)
av = av & ")"
response.write av
%>
//	
function loadModels(){
	mmake = document.theform.Make.value.replace(" ","")
	mmake = mmake.replace("-","")
	//alert(Chevrolet[1])
	var modelobj = actb(document.getElementById('Model'),'Acura');
}

//

<%
response.write "var clist = new Array()" & chr(10)
response.write "clist[0] = {Engine: 'None'}" & chr(10)
cntr = 0
j = 1
k = 1
do while cntr < 89
	if cntr >= 0 and cntr <= 9 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'1." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 10 and cntr <= 19 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'2." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 20 and cntr <= 29 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'3." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 30 and cntr <= 39 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'4." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 40 and cntr <= 49 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'5." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 50 and cntr <= 59 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'6." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 60 and cntr <= 69 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'7." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 70 and cntr <= 79 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'8." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	if cntr >= 80 and cntr <= 89 then
		for j = 1 to 10
			response.write "clist["&cntr+j&"] = {Engine:'9." & right(j,1) & " Liter'}" & chr(10)
		next
	end if
	cntr = cntr + 10
loop
%>
	function autoComplete(frmField, newFld){
		window.onerror = stopError;
		window.status = "Shop Boss Pro"
		for (i=0; i<clist.length; i++){
			if (clist[i].Engine.indexOf(frmField) != -1){
				newFld.value = clist[i].Engine
				newFld.focus()
				newFld.select()
				window.status = "Shop Boss Pro"
			}
		}
	}

function convCid(e){
	cval = e / 61.02
	cval = Math.round(cval*10)/10
	document.theform.Engine.value=cval + " LITER"
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

		//alert(document.theform.Vin.value)
	}
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

function subForm(){
	<%

	' list all required fields
	shopid = request.cookies("shopid")
	cr = chr(10) & chr(13)
	stmt = "select * from reqfields where shopid = '" & shopid & "'"
	set rs = con.execute(stmt)
	if rs("reqvin") = "yes" then 
		response.write "if(document.theform.Vin.value.length == 0){alert('" & vl & " is a required field');return;}" & cr
		vstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqyear") = "yes" then 
		response.write "if(document.theform.year.value.length == 0){alert('" & yl & " is a required field');return;}" & cr
		ystar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqmake") = "yes" then 
		response.write "if(document.theform.Make.value.length == 0){alert('" & ml & " is a required field');return;}" & cr
		mstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqmodel") = "yes" then 
		response.write "if(document.theform.Model.value.length == 0){alert('" & modl & " is a required field');return;}" & cr
		modstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqengine") = "yes" then 
		response.write "if(document.theform.Engine.value.length == 0){alert('" & el & " is a required field');return;}" & cr
		estar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqcyl") = "yes" then 
		response.write "if(document.theform.Cyl.value.length == 0){alert('" & cyll & " is a required field');return;}" & cr
		cylstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqtrans") = "yes" then 
		response.write "if(document.theform.Trans.value.length == 0){alert('" & tl & " is a required field');return;}" & cr
		tstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqdrive") = "yes" then 
		response.write "if(document.theform.drivetype.value.length == 0){alert('" & dl & " is a required field');return;}" & cr
		dstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqlicense") = "yes" then 
		response.write "if(document.theform.License.value.length == 0){alert('" & ll & " is a required field');return;}" & cr
		lstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqstate") = "yes" then 
		response.write "if(document.theform.State.value.length == 0){alert('" & sl & " is a required field');return;}" & cr
		sstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqfleet") = "yes" then 
		response.write "if(document.theform.fleetno.value.length == 0){alert('" & fl & " is a required field');return;}" & cr
		fstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqcurrmileage") = "yes" then 
		response.write "if(document.theform.currentmileage.value.length == 0){alert('" & cml & " is a required field');return;}" & cr
		cmstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if
	if rs("reqcolor") = "yes" then 
		response.write "if(document.theform.color.value.length == 0){alert('" & cl & " is a required field');return;}" & cr
		cstar = "<span style='font-size:18px;color:red;font-weight:bold'>*</span>"
	end if

	%>
	document.theform.submit()

}

function cancelCarfax(){

	document.getElementById("hider").style.display = "none"
	document.getElementById("popup").style.display = "none"

}

function carfaxLookup(){

	if (document.getElementById('carfaxvin').value.length != 17){
		if (document.getElementById('carfaxlic').value.length == 0){alert("License # is required");return}
		if (document.getElementById('carfaxlic').value.length > 0 && document.getElementById('carfaxstate').value.length == 0){alert("License Plate state is required when decoding by license number");return}
	}

	xmlHttp=getXMLHttpRequest()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	}
	trand = parseInt(Math.random()*99999999);  // cache buster
	qs = "platenumber="+document.getElementById("carfaxlic").value +"&platestate="+document.getElementById("carfaxstate").value
	qs += "&vin="+document.getElementById("carfaxvin").value
	url="carfax/quickvinrequest.asp?"+qs+"&shopid=<%=shopid%>"
	//alert(url)
	xmlHttp.onreadystatechange=carfaxstatechanged 
	xmlHttp.open("get",url,true)
	xmlHttp.send(null)



}

function carfaxstatechanged(){

	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
			rt = xmlHttp.responseText
			x = rt.indexOf("|")
			if(x > 0){
				// split to an array to populate fields
				tar = rt.split("|")

				//year,make,model,cyl,engine,trans,drive,vin,license,state
				document.theform.year.value = tar[0]
				document.theform.Make.value = tar[1]
				document.theform.Model.value = tar[2]
				document.theform.Cyl.value = tar[3]
				document.theform.Engine.value = tar[4]
				document.theform.Trans.value = tar[5]
				document.theform.drivetype.value = tar[6]
				document.theform.Vin.value = tar[7]
				document.theform.License.value = tar[8]
				document.theform.State.value = tar[9]
				//alert("Decode Complete")
				cancelCarfax()
				document.theform.color.focus()
			}else{
				alert(rt)
			}
	}
}
function getXMLHttpRequest() 
{
    if (window.XMLHttpRequest) {
        return new window.XMLHttpRequest;
    }
    else {
        try {
            return new ActiveXObject("MSXML2.XMLHTTP.3.0");
        }
        catch(ex) {
            return null;
        }
    }
}

function openCarfax(){
	document.getElementById("hider").style.display = "block"
	document.getElementById("popup").style.display = "block"

}

	var isMobile = {
	    Android: function() {
	        return navigator.userAgent.match(/Android/i);
	    },
	    BlackBerry: function() {
	        return navigator.userAgent.match(/BlackBerry/i);
	    },
	    iOS: function() {
	        return navigator.userAgent.match(/iPhone|iPad|iPod/i);
	    },
	    Opera: function() {
	        return navigator.userAgent.match(/Opera Mini/i);
	    },
	    Windows: function() {
	        return navigator.userAgent.match(/IEMobile/i);
	    },
	    any: function() {
	        return (isMobile.Android() || isMobile.BlackBerry() || isMobile.iOS() || isMobile.Opera() || isMobile.Windows());
	    }
	};
	$(document).ready(function(){
		if (isMobile.Android() || isMobile.iOS()){
			$('#scanvin').show()
			$('#scanvinwocarfax').show()
		}
	});
    function scanvincarfax(){
    	
    	if (isMobile.Android() || isMobile.iOS()){
    		bridgeit.scan('id', 'onAfterCaptureScan');
    	}
    	
    }
    
    function scanvin(){
    	if (isMobile.Android() || isMobile.iOS()){
    		bridgeit.scan('id', 'onAfterCaptureScanWOCarfax');
    	}

    }
    
    function onAfterCaptureScanWOCarfax(event){
    	var vin = event.value;
	    if (vin.length == 18){
			vin = vin.substring(1)
		}

    	document.theform.Vin.value = vin
    	$('#pbar').show()
    	showList(vin)
    	
    }
	function onAfterCaptureScan(event)  {
	    var vin = event.value;
	    if (vin.length == 18){
			vin = vin.substring(1)
		}

	    $('#carfaxvin').val(vin)
	    carfaxLookup()
	    $('#scanvinwocarfax').hide()

	}
	
	

</script>

<head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<link rel="stylesheet" type="text/css" href="css/dhtmlXCombo.css">

<style type="text/css">
<!--
.obutton{
	height:30px;
	font-family:Arial, Helvetica, sans-serif;
	font-size:small;
	font-weight:bold;
	color:maroon;
	cursor:hand;
}
obutton:hover{
	color:blue;
}

p, td, th, li { font-size: 10pt; font-family: Verdana, Arial, Helvetica }
.style1 {
	border: 1px solid #203F80;
}
.style2 {

}	
.style3 {
	font-weight: bold;
	width: 30%;
	text-align: left;
}
.style4 {
	background-color: #FFFFFF;
	width:70%

}
.style5 {
	font-size: xx-small;
}

#hider{
	position:absolute;
	top:0px;
	left:0px;
	width:100%;
	height:120%;
	background-color:gray;
	-ms-filter:"progid:DXImageTransform.Microsoft.Alpha(Opacity=70)";
	filter: alpha(opacity=70); 
	-moz-opacity:.70; 
	opacity: .7;
	z-index:998;
	display:none;

}

#popup{
	position:absolute;
	top:100px;
	left:20%;
	width:60%;
	min-height:200px;
	overflow-y:auto;
	border:medium navy outset;
	border-radius:15px;
	box-shadow:5px 5px 5px;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
	font-family:arial
}


.style6 {
	font-weight: bold;
	background-image: url('newimages/wipheader.jpg');
}
.auto-style1 {
	text-align: center;
}
.auto-style2 {
	border: 4px solid #FF0000;
	background-color: #FFFFFF;
	width: 70%;
	font-size: small;
}
.auto-style3 {
	color: #FF0000;
	font-size: medium;
}
-->
</style>
</head>
<body onload="" link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<br/><br/>
<div id="hider"></div>
<div id="popup">
<%if carfax <> "no" then%>
	As a CarFax Partner you can enter the License Number and state in the boxes 
	below. You can also enter the VIN if you like.
<%else%>
	Since you are not a CarFax partner, you will need to enter the VIN for 
	vehicle decoding or cancel to skip the VIN decode.
<%end if%>
<br><br>
<%if carfax <> "no" then%>
	License #<br>
<input id="carfaxlic" name="carfaxlic" type="text" style="width: 200px;text-transform:uppercase" ><br>
	<br>License State (2 letters only)<br>
<input id="carfaxstate" name="carfaxstate" type="text" style="width: 200px;text-transform:uppercase" ><br>
<%end if%>
	<br>VIN<br>
<input id="carfaxvin" name="carfaxvin" type="text" style="width: 200px;text-transform:uppercase" ><br>
<br>
<input id="scanvin" style="display:none;font-size:large;font-weight:bold" name="Button1" type="button" onclick="scanvincarfax()" value="Scan VIN" >
<input onclick="carfaxLookup()" name="Button1" type="button" value="Decode Vehicle" > <input onclick="cancelCarfax()" name="Button1" type="button" value="Close" >
<br><br>
	Please Note: Carfax does not have data on every vehicle. If the License # 
	you entered is not found, please enter the VIN here or close this window to 
	use our Classic VIN Lookup function.
</div>
<div align="center">
 <center>
 <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
  <tr>
   <td width="100%" valign="top">
    <div align="left">
     <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
      <tr>
       <td valign="top" width="100%">
        <div align="center">
         <table border="0" cellpadding="0" cellspacing="0" width="70%">
          <tr>
           <td height="11" class="style6">
            <p style="width:100%;padding:10px;font-size:medium" align="center"><font color="#FFFFFF">
			Please enter vehicle information, then click Add Vehicle</font><br>
              </td>
          </tr>
         </table>
        </div>
        <div align="center">
         <center>
         <%
         if android = "yes" then
         %>
         <input id="scanvinwocarfax" type="button" tabindex="15" onclick="location.href='zxing://scan/?ret=http://<%=request.servervariables("SERVER_NAME")%>/sbp/addveh-new.asp?vin={CODE}|<%=request("custid")%>|customer2.asp'" style="width:300px;height:50px;font-weight:bold;font-size:x-large" name="Abutton1" value="Click to Scan VIN">
		 <%
		 end if
		 %>
         <form name="theform" method="post" action="addvehnow.asp">
         <input type="hidden" name="origshopid" value="<%=request("origshopid")%>">
         <input type="hidden" name="reason" value="<%=request("reason")%>">
          <table cellspacing="0" cellpadding="2" class="style1" style="width: 70%">
           <tr>

            <td class="style3" style="width: 20%"><%=vl & vstar%></td>
            <td  class="style4" style="width: 60%">
			<input onblur="javascript:document.theform.Vin.style.backgroundColor='white'" onfocus="javascript:document.theform.Vin.style.backgroundColor='#FFFF99'" type="text" name="Vin" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Vin")%>">
			
			<input type="button" tabindex="15" onclick="pdmv('c')" name="Abutton1" value="Click to Decode">&nbsp;</td>
            <td  class="auto-style2" style="width: 60%;padding:10px;" rowspan="21" valign="top">
			<span class="auto-style3"><strong>*** NOTE ***</strong></span><strong><br>
			<br>Vehicle fields can now be customized <em>AND </em>marked as 
			required or not required.<br><br><br>To customize a label for a 
			vehicle field, go to Settings -&gt; Custom.&nbsp; There you can change the 
			name of any field and mark them as required or not required.<br><br><br>
			You can also add additional custom vehicle fields to save 
			information other than the standard fields.</strong></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=yl & ystar%></td>
            <td  class="style4" style="width: 60%">
			<input onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="year" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Year")%>">
				<div style="padding-left:4px;padding-right:4px;padding-top:2px;height:22px;background-color:black;margin-top:-22px;width:250px;display:none" id="pbar">
					<font color="red"><b>Decoding VIN. Please stand by...</b></font></div></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=ml & mstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="Make" onblur="this.style.backgroundColor='white';" onfocus="javascript:loadMakes(this,mlist,document.theform.Model);this.style.backgroundColor='#FFFF99'" type="text" name="Make" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray; width: 157px;" value='<%=request("make")%>'></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=modl & modstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="Model" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:loadMakes(this,document.theform.Make.value,document.theform.Engine);this.style.backgroundColor='#FFFF99'" type="text" name="Model" size="20" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("model")%>'>
			</td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=cyll & cylstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="cyl"  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:loadMakes(this,document.theform.Make.value,document.theform.Engine);this.style.backgroundColor='#FFFF99'" type="text" name="Cyl" size="20" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("cyl")%>'></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=el & estar%></td>
            <td  class="style4" style="width: 60%">
			<input onfocus="this.style.backgroundColor='#FFFF99'" onblur="this.style.backgroundColor='white'" onkeyup="javascript:if (this.value.length == 3){autoComplete(this.value, document.theform.Engine)};" type="text" name="Engine" size="14" value="<%=request("Engine")%>" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; border-style: solid; border-color: gray">
             (ie. 2.4 Liter)</td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=cl & cstar%></td>
            <td  class="style4" style="width: 60%">
			<input id="color"  onblur="javascript:document.theform.Cyl.style.backgroundColor='white'" onfocus="javascript:document.theform.Cyl.style.backgroundColor='#FFFF99';" type="text" name="Color" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Cyl")%>"></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=tl & tstar%></td>
            <td  class="style4" style="width: 60%">
			<input  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="Trans" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("Trans")%>"> 
			(AUTO, MAN, 5SP)</td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%; height: 24px;"><%=dl & dstar%></td>
            <td  class="style4" style="width: 60%; height: 24px;">
			<input onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFF99'" type="text" name="drivetype" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value='<%=request("fleetno")%>'> 
			(RWD, FWD, AWD)</td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=ll & lstar%></td>
            <td  class="style4" style="width: 60%">
			<input  onblur="javascript:document.theform.License.style.backgroundColor='white'" onfocus="javascript:document.theform.License.style.backgroundColor='#FFFF99'" type="text" name="License" size="14" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" value="<%=request("License")%>"></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=sl & sstar%></td>
            <td  class="style4" style="width: 60%">
			<input onblur="this.style.backgroundColor='white'" onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" type="text" name="State" size="6" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray; background-color:white" value="<%=shopstate%>"></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=fl & fstar%></td>
            <td  class="style4" style="width: 60%">
			<input  
			onblur="javascript:document.theform.State.style.backgroundColor='white'" 
			onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" 
			type="text" name="fleetno" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value='<%=request("fleetno")%>'></td>
           </tr>
           <tr>
            <td class="style3" style="width: 20%"><%=cml & cmstar%></td>
            <td  class="style4" style="width: 60%">
			<input  
			onblur="javascript:document.theform.State.style.backgroundColor='white'" 
			onfocus="javascript:document.theform.State.style.backgroundColor='#FFFF99'" 
			type="text" name="currentmileage" size="20" 
			style="font-family: Arial; font-size: 10pt; font-variant: small-caps; border-style: solid; border-color: gray" 
			value='<%=request("currentmileage")%>'></td>
           </tr>
           <%	if len(vf1) > 0 then	%>
           <tr>
            <td class="style3" style="width: 20%"><%=vf1%></td>
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
            <td class="style3" style="width: 20%"><%=vf2%></td>
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
            <td class="style3" style="width: 20%"><%=vf3%></td>
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
            <td class="style3" style="width: 20%"><%=vf4%></td>
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
            <td class="style3" style="width: 20%"><%=vf5%></td>
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
            <td class="style3" style="width: 20%"><%=vf6%></td>
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
            <td class="style3" style="width: 20%"><%=vf7%></td>
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
            <td class="style3" style="width: 20%"><%=vf8%></td>
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
				<td colspan="3" class="auto-style1"><strong><br>Need more 
				fields? Go to Settings -&gt; Custom. There you can add up to 8 
				additional custom fields</strong></td>
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
    </div>
   </td>
  </tr>
 </table>
 </center>
</div>

<script language="javascript">

$(document).ready(function(){
	//alert(document.theform.Vin.value.length)
	if (document.theform.Vin.value.length == 17){
		//alert("decoding")
		decodeVin(document.theform.Vin.value)
	}


});


darray = new Array("FWD","RWD","4WD","AWD","DRWD","D4WD")

function loadDT(e,a,nf){
	if(a != darray){
		a = a.replace(" ","")
		a = a.replace("-","")
		aa = eval(a)
	}else{
		aa = a
	}
	var mobj = actb(e,aa,nf);

}

	function loadMakes(e,a,nf){
	if(a != mlist){
		a = a.replace(" ","")
		a = a.replace("-","")
		aa = eval(a)
	}else{
		aa = a
	}
	var mobj = actb(e,aa,nf);
	}

	document.theform.Vin.focus()
</script>

<!-- #include file=helpfiles/addvehhelp.asp -->
</html>
<%
'Copyright 2011 - Boss Software Inc.
%>