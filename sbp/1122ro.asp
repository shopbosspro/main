<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->


<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<style type="text/css">
<!--
a            { text-decoration: none }
-->
select{
	width:150px;
}
   code, pre {font-family: Consolas,monospace; color: green;}
   ol li {margin: 0 0 15px 0;}

</style>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<%
on error resume next
function formatdollar(a)

	if isnull(a) then
		formatdollar = "0.00"
	else	
		formatdollar = formatnumber(a,2)
	end if

end function
'check for required fields
strsql = "Select * from repairorders where shopid = '" & request.cookies("shopid") & "' and ROID = " & request("roid")
set rs = con.execute(strsql)
rodisc = rs("rodisc")
warrdisc = rs("warrdisc")
cid = rs("customerid")
if len(rs("tagnumber")) > 0 then
	roid = rs("tagnumber")
else
	roid = rs("roid")
end if
'calculate gp
pc = rs("partscost")
if rs("datein") >= #11/11/2005# then
	pro = "printronewall.asp"
else
	pro = "printro.asp"
end if
if ucase(rs("Status")) = "CLOSED" then 
	sstat = "closed"
	if request("sp") = "history.asp" then
		newsp = "viewro.asp?lic=" & request("lic") & "&roid=" & request("roid") & "&sp=" & request("sp")
	elseif request("sp") = "findroframe.asp" then
		newsp = "viewro.asp?roid=" & request("roid") & "&srch=" & request("srch") & "&sf=" & request("sf") & "&SortBy=" & request("SortBy") & "&sp=" & request("sp")
	elseif request("sp") = "wip.asp" then
		newsp = "viewro.asp?roid=" & request("roid")
	elseif request("sp") = "viewro.asp" then
		newsp = "viewro.asp?roid=" & request("roid")
	else
		newsp = "wip.asp"
	end if
	response.redirect newsp
end if
rstmt = "Select * from company where shopid = '" & request.cookies("shopid") & "'"
set rrs = con.execute(rstmt)
rrs.movefirst
userfee1max = rrs("userfee1max")
userfee2max = rrs("userfee2max")
userfee3max = rrs("userfee3max")
milesinlabel = rrs("milesinlabel")
milesoutlabel = rrs("milesoutlabel")
response.write chr(10) & "<script language=""javascript"">" & chr(10)
response.write vbTab & "function saveAll(dir, itype, sid, pid, lid){" & chr(10)
response.write vbTab & "document.getElementById('hider').style.display='block'" & chr(10)
response.write vbTab & "document.getElementById('timer').style.display='block'" & chr(10)
response.write vbTab & "document.theform.dir.value = dir" & chr(10)
response.write vbTab & "document.theform.itype.value = itype" & chr(10)
response.write vbTab & "document.theform.sid.value = sid" & chr(10)
response.write vbTab & "document.theform.pid.value = pid" & chr(10)
response.write vbTab & "document.theform.lid.value = lid" & chr(10)
response.write vbTab & "//check for miles, writer and source" & chr(10)
response.write vbTab & "i = 0" & chr(10)
response.write vbTab & "newline = '\n \n'" & chr(10)
response.write vbTab & "errmess = 'Required Fields are not Complete'" & chr(10)
rrs.movefirst

if ucase(rs("Status")) = "CLOSED" or ucase(rs("Status")) = "FINAL" then
	if rrs("RequireSource") = "yes" then
		response.write vbTab & "if (document.theform.Source.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Source is a required field.'}" & chr(10)
	end if
	if rrs("RequireComplaint") = "yes" then
		response.write vbTab & "if (document.theform.Comments.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Comments is a required field.'}" & chr(10)
	end if
	if lcase(rrs("requireoutmileage")) = "yes" then
		response.write vbTab & "if (document.theform.MilesOut.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Vehicle In Miles is required'}" & chr(10)
	end if
	if lcase(rrs("requireoutmileage")) = "yes" then
		response.write vbTab & "if (document.theform.VehicleMiles.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Vehicle Out Miles is required'}" & chr(10)
	end if
end if
response.write vbTab & "if (i > 0){" & chr(10)
response.write vbTab & vbTab & "alert(errmess+newline)" & chr(10)
response.write vbTab & "}else{calcPercents();document.theform.submit();" & chr(10)
response.write vbTab & "}setTimeout('closeTimer()',500)}" & chr(10)
response.write "</script>"
set rrs = nothing
homephone = "(" & left(rs("CustomerPhone"),3) & ") " & mid(rs("CustomerPhone"),4,3) & "-" & right(rs("CustomerPhone"),4)
workphone = "(" & left(rs("CustomerWork"),3) & ") " & mid(rs("CustomerWork"),4,3) & "-" & right(rs("CustomerWork"),4)
cellphone = "(" & left(rs("CellPhone"),3) & ") " & mid(rs("CellPhone"),4,3) & "-" & right(rs("CellPhone"),4)
'ttlsub = rs("totalsublet")
'response.write ttlsub
ttlprts = rs("totalprts")
'ttllbr = rs("totallbr")

%>

<script type="text/javascript" src="javascripts/ro.js"></script>
<link rel="STYLESHEET" type="text/css" href="css/rich_calendar.css">
<script language="JavaScript" type="text/javascript" src="javascripts/rich_calendar.js"></script>
<script language="JavaScript" type="text/javascript" src="javascripts/rc_lang_en.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/domready.js"></script>
<script type="text/javascript">
var cal_obj2 = null;

var format = '%m/%d/%Y';

// show calendar
function show_cal(el,a) {

	if (cal_obj2) return;
	document.getElementById("dfield").value = a
	var text_field = document.getElementById(document.getElementById("dfield").value);
	text_field.disabled = true;
	cal_obj2 = new RichCalendar();
	cal_obj2.start_week_day = 0;
	cal_obj2.show_time = false;
	cal_obj2.language = 'en';
	cal_obj2.user_onchange_handler = cal2_on_change;
	cal_obj2.user_onclose_handler = cal2_on_close;
	cal_obj2.user_onautoclose_handler = cal2_on_autoclose;

	cal_obj2.parse_date(text_field.value, format);

	cal_obj2.show_at_element(text_field, "adj_right-");
	//cal_obj2.change_skin('alt');
	

}

// user defined onchange handler
function cal2_on_change(cal, object_code) {
	if (object_code == 'day') {
		document.getElementById(document.getElementById("dfield").value).disabled = false;
		document.getElementById(document.getElementById("dfield").value).value = cal.get_formatted_date(format);
		//alert(document.getElementById("caldatefield").value)
		if(document.getElementById("caldatefield").value == "StatusDate" || document.getElementById("caldatefield").value == "DateIn"){
			chgDB(document.getElementById("caldatefield").value)
		}
		//alert(document.theform.dbStatusDate.value+":"+document.theform.dbDateIn.value)
		cal.hide();
		cal_obj2 = null;

	}
}

// user defined onclose handler (used in pop-up mode - when auto_close is true)
function cal2_on_close(cal) {
		document.getElementById(document.getElementById("dfield").value).disabled = false;
		cal.hide();
		cal_obj2 = null;

}

// user defined onautoclose handler
function cal2_on_autoclose(cal) {
	document.getElementById(document.getElementById("dfield").value).disabled = false;
	cal_obj2 = null;
}
function showInspection(){

	var r=Math.floor(Math.random()*100001)
	document.getElementById("inspection").src = "inspection/inspection.asp?roid=<%=request("roid")%>&r="+r 
	document.getElementById("inspection").style.display = "block"
	document.getElementById("popuphider").style.display = "block"

}

</script>
<style type="text/css">
<!--
p, td, th, li { font-size: 10pt; font-family: Verdana, Arial, Helvetica }
.style1 {
	border-width: 0;
	font-family: Arial;
	color: white;
	background-color: #0066CC;
}
.abuts{
	font-size:12px;
	width:100px;
	height:24px;
	cursor:hand;
	color:black;
	
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
#timer{
	background-color:white;
	position:absolute;
	top:25%;
	left:35%;
	width:30%;
	height:20%;
	z-index:999;
	display:none;
	text-align:center;
	border:4px navy outset;
	color:black;
	padding-top:2%;
	font-family:Verdana, Arial, sans-serif;
	font-weight:bold;
}

#estimator{
	background-color:white;
	position:absolute;
	top:1%;
	left:1%;
	width:97%;
	height:95%;
	z-index:999;
	display:none;
	border:4px navy outset;
	color:black;
	padding:5px;
	overflow:auto;

}

#estimatorqueue{
	background-color:white;
	position:absolute;
	top:1%;
	left:1%;
	width:48%;
	height:16%;
	z-index:999;
	display:none;
	border:4px navy outset;
	color:black;
	padding:5px;
	overflow:auto;

}
#techselect{
	background-color:white;
	position:absolute;
	top:20%;
	left:30%;
	width:40%;
	height:25%;
	z-index:1000;
	display:none;
	border:4px navy outset;
	color:black;
	padding:5px;
	padding-top:30px;
	overflow:auto;

}

#pestimatorqueue{
	background-color:white;
	position:absolute;
	top:1%;
	right:1%;
	width:48%;
	height:16%;
	z-index:999;
	display:none;
	border:4px navy outset;
	color:black;
	padding:5px;
	overflow:auto;

}


#history{
	width:90%;
	height:90%;
	top:5%;
	left:5%;
	background-color:white;
	position:absolute;
	z-index:1000;
	border:medium navy outset;
}
-->

.menustyle{
	text-align: center;
	background-color: #F0F0F0;
	width:9%;
	border:1px gray outset;
	color:maroon;
	font-weight:bold;
	cursor:pointer;
}

.style8 {
	text-align: center;
	color: #FFFFFF;
	border-width: 1px;
	background-image:url('newimages/pageheader.jpg');
}

.style9 {
	color: #FFFFFF;
	font-weight: bold;
	border-width: 1px;
	background-image:url('newimages/rosubheader.jpg');
}

.style10 {
	background-image: url('newimages/wipheader.jpg');
	color:white;
	font-weight:bold;
}

.style13 {
	color: #FFFFFF;
}
.style14 {
	color: white;
	border-width: 0;
	background-color: #0066CC;
}
.style15 {
	border-width: 0;
}

.style16 {
	border-collapse: collapse;
	border-style: solid;
	border-width: 1px;
}
.style17 {
	border-style: solid;
	border-width: 0;
	color: #FFFFFF;
	background-color: #0066CC;
}
.style18 {
	text-align: right;
}
.style20 {
	border-width: 0;
	text-align: right;
}
.style21 {
	color: #FFFFFF;
	border-width: 1px;
	background-color: #0066CC;
}
p,div,span,body{
	font-family:Arial, Helvetica, sans-serif
}

#popup{
	position:absolute;
	top:100px;
	left:40%;
	width:400px;
	min-height:200px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
}
#popuphider{
	position:absolute;
	top:0px;
	left:0px;
	width:100%;
	height:100%;
	background-color:gray;
	-ms-filter:"progid:DXImageTransform.Microsoft.Alpha(Opacity=50)";
	filter: alpha(opacity=70); 
	-moz-opacity:.70; 
	opacity: .7;
	z-index:997;
	display:none;

}

.style22 {
	font-size: x-small;
}

#recrepairs{
	position:absolute;
	top:100px;
	left:10%;
	width:80%;
	height:500px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:1050;
	background-color:white;
	padding:20px;
	font-size:medium;
	font-weight:bold;

}
#addrecrepairs{
	position:absolute;
	top:100px;
	left:15%;
	width:60%;
	height:500px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
	font-size:medium;
	font-weight:bold;

}
#editrecrepairs{
	position:absolute;
	top:100px;
	left:15%;
	width:60%;
	height:500px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
	font-size:medium;
	font-weight:bold;

}

#disclosure{
	position:absolute;
	top:100px;
	left:15%;
	width:60%;
	height:500px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
	font-size:medium;
	font-weight:bold;

}
#inspection{
	width:90%;
	height:90%;
	top:5%;
	left:5%;
	background-color:white;
	position:absolute;
	z-index:1000;
	border:medium navy outset;
}

#customvehiclefields{
	position:absolute;
	top:100px;
	left:30%;
	width:40%;
	height:500px;
	overflow-y:auto;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
	padding:20px;
	font-size:medium;
	font-weight:bold;
	text-align:left

	
}

.style24 {
	border-style: solid;
	border-width: 0;
	text-align: center;
	color: #FFFFFF;
	background-color: #006600;
}
.style25 {
	border-style: solid;
	border-width: 0;
	color: #FFFFFF;
	background-color: #0066CC;
	text-align: center;
}
.style26 {
	text-align: center;
}
</style>

<script type="text/javascript" language="javascript">



function closeTimer(){

	
		document.getElementById('timer').style.display = 'none'
		document.getElementById('hider').style.display = 'none'



}

var xmlHttp
function showHistory(){
	document.getElementById("hider").style.display = "block"

	document.getElementById("history").style.display = "block"

}

function reloadPage(){
	document.body.onunload='';
	saveAll('ro.asp?a=n','','','','','')
}

function chgDB(fldchg){
	var dbchg = "db"+fldchg
	//alert(dbchg)
	docitem = "document.theform."+fldchg
	//alert(docitem)
	docval = eval(docitem+".value")
	//alert(docval)
	newdocval = replace(replace(docval,'\r',' '),'\n',' ')
	fldval = "document.theform." + dbchg + ".value='" + newdocval + "'"
	//alert(fldval)
	eval(fldval)
}
function closeRO(){
	if (document.theform.Status.value=='CLOSED'){
		if (confirm('This will close the RO.  Click OK to proceed or Cancel to Cancel')!=0){
			document.theform.Status.focus();
			var date=new Date();
			var currentTime = new Date()
			var month = currentTime.getMonth() + 1
			var day = currentTime.getDate()
			var year = currentTime.getFullYear()
			var finaldate = month+"/"+day+"/"+year
			document.theform.dbFinalDate.value = finaldate
			document.theform.dbStatusDate.value = finaldate
			saveAll('wip.asp?a=n','','','','','')
		}else{
			document.theform.Status.value='Final'
		}
	}
}
function replace(string,text,by) {
// Replaces text with by in string
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return string;

    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return string;
    if (i == -1) return string;

    var newstr = string.substring(0,i) + by;

    if (i+txtLength < strLength)
        newstr += replace(string.substring(i+txtLength,strLength),text,by);

    return newstr;
}
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strMonth=dtStr.substring(0,pos1)
	var strDay=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		//alert("The date format should be : mm/dd/yyyy")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		//alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		//alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		//alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		//alert("Please enter a valid date")
		return false
	}
return true
}

function ValidateForm(){
	var dt=document.frmSample.txtDate
	if (isDate(dt.value)==false){
		dt.focus()
		return false
	}
    return true
 }


function setStatusDate(){
	var date=new Date();
	var d=date.getDate();
	var day = (d < 10) ? '0' + d : d;
	var m = date.getMonth() + 1;
	var month = (m < 10) ? '0' + m : m;
	var yy = date.getYear();
	var year = (yy < 1000) ? yy + 1900 : yy;
	var finaldate = month+"/"+day+"/"+year
	if(document.theform.Status.value=="FINAL" || document.theform.Status.value=="CLOSED"){
		
		if(document.theform.dbFinalDate.value.length > 0){
			//do nothing
		}else{
		//alert("first loop")
		document.theform.dbFinalDate.value = finaldate
		document.theform.dbStatusDate.value = finaldate
		}
	}else{
		document.theform.dbStatusDate.value = finaldate
	}

}

function changeColor(a){

	if(document.getElementById(a).style.backgroundColor = "white"){
		document.getElementById(a).style.backgroundColor = "green"
		document.getElementById(a).style.backgroundColor = "white"
	}else{
		document.getElementById(a).style.backgroundColor = "white"
		document.getElementById(a).style.backgroundColor = "#000080"
	}

}

function uploadFiles(){

	document.getElementById("showframe").style.display = "block"
	document.getElementById("hider").style.display = "block"


}

function showPics(){

	document.getElementById("picframe").style.display = "block"
	document.getElementById("hider").style.display = "block"


}

function closeUpdateWindow(){
	
	if(document.theform.x.value == "y"){
		document.getElementById("hider").style.display = "none"
		document.getElementById("alerts").style.display = "none"
		document.theform.x.value = ""
		return;
	}
	t=setTimeout('document.theform.submit()',500)

}

function showUpdateWindow(x){
	
	if (x == "y"){
		document.getElementById("hider").style.display = "block"
	
		document.getElementById("alerts").style.display = "block"
		document.theform.x.value = "y"
		return;
	}
	//alert(document.theform.Status.value +":"+ document.theform.oldstatus.value)
	if(document.theform.Status.value != document.theform.oldstatus.value){
		document.getElementById("hider").style.display = "block"
	
		document.getElementById("alerts").style.display = "block"

	}else{
		t=setTimeout('document.theform.submit()',500)
	}


}
var xmlHttp

function stateChanged(){ 
	//alert(xmlHttp.readyState)
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
		//alert("Update Sent")
		alert(xmlHttp.responseText)
		closeUpdateWindow()
		
	} 
} 

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

function sendUpdate(t){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="rosendupdate.asp?t="+t+"&roid=<%=request("roid")%>"
	xmlHttp.onreadystatechange=stateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function printtopdf(){

	window.onunload = ""
	document.location.href='printropdf.asp?roid=<%=request("roid")%>'

}

function closeGuide(){

	document.getElementById("estimator").style.display = "none"
	document.getElementById("hider").style.display = "none"
	document.getElementById("estimatorqueue").style.display = "none"
	document.getElementById("pestimatorqueue").style.display = "none"


}

var r, c

function showGuide(roid,comid,v){

	document.getElementById("estimator").innerHTML = '<img style="position:absolute;top:40%;left:45%" src="newimages/ajax-loader.gif">'
	r = roid
	c = comid
	document.getElementById("estimator").style.display = "block"
	document.getElementById("hider").style.display = "block"

	document.guideform.comid.value = comid
	
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	if (v == "y"){
		if (vin.length == 17){
			alldataurl = "vins/"+vin
		}else{
			alldataurl = "years"
		}
	}
	if (v == "n"){
		alldataurl = "years"
	}
	//alert(alldataurl)
	var url="guide/alldatadisplay.asp?roid=<%=request("roid")%>&comid="+comid+"&r="+r+"&alldatashopid=<%=request.cookies("shopid")%>&alldataurl="+alldataurl
	//alert(url)
	xmlHttp.onreadystatechange=guideStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function guideStateChanged(){

	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		rt = xmlHttp.responseText
		if (rt == "bad vin"){
			sg = "showGuide('"+r+"','"+c+"','n')"
			eval(sg)
			document.getElementById("estimator").innerHTML = message
			
		}else{
			document.getElementById("estimator").innerHTML = rt
		}
		//alert(xmlHttp.responseText)
	}

}



function getLink(shopid,url,roid,comid){

	document.getElementById("estimator").innerHTML = '<img style="position:absolute;top:40%;left:45%" src="newimages/ajax-loader.gif"></span>'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	var url="guide/alldatadisplay.asp?roid=<%=request("roid")%>&comid="+document.guideform.comid.value+"&r="+r+"&alldatashopid=<%=request.cookies("shopid")%>&alldataurl="+url
	//alert(url)
	xmlHttp.onreadystatechange=guideStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function selectTech(){

	document.getElementById("techselect").style.display = "block"

}

function saveTech(tech,descriptionName){

	if(t != "select"){
		document.getElementById("labor*"+descriptionName+"*de").value = tech
	}else{
		alert("Please select a technician for this labor operation")
	}

}

function addLabor(operationName,componentName,componentItem,laborTime){

	//create a tech dropdown for the labor added queue
	<%
	techbuild = "<select style='font-size:10px;' size='1' id='tech' name='tech'>"
	strsql = "Select EmployeeLast, EmployeeFirst from employees where shopid = '" & request.cookies("shopid") & "' and Active = 'yes'"
	set techrs = con.execute(strsql)
	
	if not techrs.eof then
	techrs.movefirst
	while not techrs.eof
		techbuild = techbuild & "<option value='" & techrs ("EmployeeLast") & ", " & techrs ("EmployeeFirst")
		techbuild = techbuild & "'>" & techrs ("EmployeeLast") & ", " & techrs ("EmployeeFirst") & "</option>"
		techrs.movenext
	wend
	end if
	set techrs = nothing
	techbuild = techbuild & "</select>"
	response.write "techbuild = """ & techbuild & """"
	%>

	//set new line to be added
	descriptionname = removeSpaces(operationName+componentName+componentItem)
	timename = removeSpaces(operationName+componentName+componentItem+laborTime)
	newline = "<span style='font-size:12px;font-weight:bold;'>Desc:</span> <input style='font-size:10px;' size='30' type='text' id='labor*"+descriptionname+"*de' name='labor*"+descriptionname+"*de' value='"+operationName+" " +componentName+" "+componentItem+"'>"
	newline += " <span style='font-size:12px;font-weight:bold;'>Hours:</span> <input style='font-size:10px;' id='labor*"+descriptionname+"*ti' name='labor*"+descriptionname+"*ti' type='text' size='10' value='"+laborTime+"'>"
	techbuild = techbuild.replace(/tech/gi,"labor*"+descriptionname+"*tech")
	newline += " <span style='font-size:12px;font-weight:bold;'>Tech: </span>" + techbuild
	newline += "<input style='font-size:12px' id='labor*"+descriptionname+"*button' type='button' value='Remove' onclick='deleteLaborItem(\""+descriptionname+"\")'><br id='labor*"+descriptionname+"br'>"
	document.getElementById("estimator").style.top = "19%"
	document.getElementById("estimator").style.height = "80%"
	document.getElementById("estimatorqueue").style.display = "block"
	document.getElementById("estimatorqueue").innerHTML += newline
	document.guideform.laboritems.value += descriptionname + "~"
	//alert(document.getElementById("estimatorqueue").innerHTML)
	//alert(document.getElementById("labor*"+descriptionname+"*tech")
	

}

function addPart(pPartDescription,pPartNumber,pQuantity,pPartPrice){

	newline = "<span style='font-size:12px;font-weight:bold;'>#: </span><input style='font-size:10px;' name='part*"+pPartNumber+"*pn' size='20' type='text' name='"+pPartNumber+"pn' value='"+pPartNumber+"'>"
	newline += "<span style='font-size:12px;font-weight:bold;'> Desc: </span><input style='font-size:10px;' name='part*"+pPartNumber+"*de' size='50' type='text' name='"+pPartDescription+"pn' value='"+pPartDescription+"'>"
	newline += "<span style='font-size:12px;font-weight:bold;'> Qty: </span><input style='font-size:10px;' name='part*"+pPartNumber+"*qty' type='text' size='4' value='"+pQuantity+"'>"
	newline += "<span style='font-size:12px;font-weight:bold;'> Price: </span><input size='10' style='font-size:10px' name='part*"+pPartNumber+"*pr' type='text' value='"+pPartPrice+"'>"
	newline += " <input style='font-size:10px' name='part*"+pPartNumber+"*button' type='button' value='Remove' onclick='deletePartItem(\""+pPartNumber+"\")'><br id='part*"+pPartNumber+"*br'>"
	document.getElementById("estimator").style.top = "19%"
	document.getElementById("estimator").style.height = "80%"
	document.getElementById("pestimatorqueue").style.display = "block"
	document.getElementById("pestimatorqueue").innerHTML += newline
	document.getElementById("estimatorqueue").style.display = "block"
	document.guideform.partitems.value += pPartNumber + "~"


}

function removeSpaces(string) {
 return string.split(' ').join('');
}


function deletePartItem(n){
	
	nelement = document.getElementById("part*"+n+"*pn")
	velement = document.getElementById("part*"+n+"*de")
	belement = document.getElementById("part*"+n+"*button")
	brelement = document.getElementById("part*"+n+"*br")
	qelement = document.getElementById("part*"+n+"*qty")
	pelement = document.getElementById("part*"+n+"*pr")
	nelement.style.display = "none"
	velement.style.display = "none"
	belement.style.display = "none"
	brelement.style.display = "none"
	qelement.style.display = "none"
	pelement.style.display = "none"
	nelement.parentNode.removeChild(nelement)
	velement.parentNode.removeChild(velement)
	belement.parentNode.removeChild(belement)
	brelement.parentNode.removeChild(brelement)
	qelement.parentNode.removeChild(qelement)
	pelement.parentNode.removeChild(pelement)
	document.guideform.partitems.value.replace(n+"~","")
	//alert(document.getElementById("estimatorqueue").innerHTML)

}


function deleteLaborItem(n){

	delement = document.getElementById("labor*"+n+"*de")
	nelement = document.getElementById("labor*"+n+"*ti")
	velement = document.getElementById("labor*"+n+"*tech")
	belement = document.getElementById("labor*"+n+"*button")
	brelement = document.getElementById("labor*"+n+"br")
	delement.style.display = "none"
	nelement.style.display = "none"
	velement.style.display = "none"
	belement.style.display = "none"
	brelement.style.display = "none"
	delement.parentNode.removeChild(delement)
	nelement.parentNode.removeChild(nelement)
	velement.parentNode.removeChild(velement)
	belement.parentNode.removeChild(belement)
	brelement.parentNode.removeChild(brelement)
	document.guideform.laboritems.value.replace(n+"~","")
	//alert(document.getElementById("estimatorqueue").innerHTML)

}

function showPayment(){
	document.getElementById("popup").style.display = "block"
	document.getElementById("popuphider").style.display = "block"


}
function closePayment(){

	document.getElementById("popup").style.display = "none"
	document.getElementById("popuphider").style.display = "none"

}

function savePayment(){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	var url="ropostpayment.asp?roid="+document.pmtform.roid.value+"&amt="+document.pmtform.amt.value+"&type="+document.pmtform.type.value+"&cid="+document.pmtform.cid.value+"&number="+document.pmtform.number.value+"&shopid=<%=request.cookies("shopid")%>"
	xmlHttp.onreadystatechange=paymentStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}

function paymentStateChanged(){

	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		rt = xmlHttp.responseText
		if(rt == "posted"){
			document.body.onunload = ""
			saveAll("ro.asp?",'','','','')
		}else{
			alert("Error Saving Payment")
		}
	}

}

function deletePayment(id){

	c = confirm("This will delete this payment.  Are you sure?")
	if(c){
		document.body.onunload = ""
		saveAll("rodeletepayment.asp?roid=<%=request("roid")%>&id="+id+"&",'','','','')
	}

}

function clearPercentFee(userfee){

	document.getElementById(userfee).value = "zero"

}

function showImage(s){
	alert(s)
	window.open(s,"image")
	
}

function clearFee(n){
	c = confirm("This will set this fees percentage to 0.  This cannot be reversed on this RO.  Are you sure?")
	if(c){
		feetoclear = eval("document.theform.UserFee" + n)
		percenttoclear = eval("document.theform.dbuserfee" + n + "percent")
		dbfeetoclear = eval("document.theform.dbUserFee" + n)
		
		feetoclear.value = '0'
		percenttoclear.value = '0'
		dbfeetoclear.value = '0'
		
		//alert(eval("document.theform.dbUserFee"+n+".value"))
		
		calcPercents()
	}
	

}

function showRecRepairs(){

	document.getElementById("recrepairs").style.display = "block"
	document.getElementById("popuphider").style.display = "block"
}

function closeRecRepairs(){

	document.getElementById("recrepairs").style.display = "none"
	document.getElementById("popuphider").style.display = "none"


}
function showVehicleFields(){
	document.getElementById("customvehiclefields").style.display = "block"
	document.getElementById("popuphider").style.display = "block"

}

function hideVehicleFields(){
	document.getElementById("customvehiclefields").style.display = "none"
	document.getElementById("popuphider").style.display = "none"

}

function showDisclosure(){
	document.getElementById("disclosure").style.display = "block"
	document.getElementById("popuphider").style.display = "block"

}

function hideDisclosure(){
	document.getElementById("disclosure").style.display = "none"
	document.getElementById("popuphider").style.display = "none"

}


function calcPercents(){

	// show max values from database
	userfee1max = '<%=userfee1max%>'
	userfee2max = '<%=userfee2max%>'
	userfee3max = '<%=userfee3max%>'

	//check for percents in the dbfields
	if(document.theform.UserFee1){
		fee1 = document.theform.UserFee1.value
		if(!isNaN(document.theform.dbuserfee1percent.value)){
			if(document.theform.dbuserfee1percent.value > 0){
				fee1val = (document.theform.dbuserfee1percent.value * document.theform.dbSubtotal.value)/100
				if(userfee1max > 0 && userfee1max < fee1val){fee1val = userfee1max}
				document.theform.UserFee1.value = fee1val
				document.theform.dbUserFee1.value = fee1val
				fee1 = fee1val

			}
		}
	}
	if(document.theform.UserFee2){
		fee2 = document.theform.UserFee2.value
		if(!isNaN(document.theform.dbuserfee2percent.value)){
			if(document.theform.dbuserfee2percent.value > 0){
				fee2val = (document.theform.dbuserfee2percent.value * document.theform.dbSubtotal.value)/100
				//alert(userfee2max+":"+fee2val)
				if(userfee2max > 0 && userfee2max < fee2val){fee2val = userfee2max}
				document.theform.UserFee2.value = fee2val
				document.theform.dbUserFee2.value = fee2val
				fee2 = fee2val
			}
		}
	}
	if(document.theform.UserFee3){
		fee3 = document.theform.UserFee3.value
		if(!isNaN(document.theform.dbuserfee3percent.value)){
			if(document.theform.dbuserfee3percent.value > 0){
				fee3val = (document.theform.dbuserfee3percent.value * document.theform.dbSubtotal.value)/100
				if(userfee3max > 0 && userfee3max < fee1val){fee1val = userfee3max}
				document.theform.UserFee3.value = fee3val
				document.theform.dbUserFee3.value = fee3val
				fee3 = fee3val
			}
		}
	}

	//alert(fee1+":"+fee2+":"+fee3)
	hw = document.theform.HazardousWaste.value
	sf = document.theform.storagefee.value
	document.theform.dbTotalFees.value = parseFloat(fee1)+parseFloat(fee2)+parseFloat(fee3)+parseFloat(hw)+parseFloat(sf)
}

function declineAcceptIssue(issueid){

	saveAll("roissuedeclineaccept.asp?comid="+issueid+"&")
	
}

function hideDrops(dropid){

	document.getElementById(dropid).style.display='none'

}

function deleteRec(id){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	var url="rodeleterec.asp?id="+id
	//alert(url)
	xmlHttp.onreadystatechange=recChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}

function recChanged(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		x = xmlHttp.responseText
		if(x == "success"){
			document.body.onunload = ""
			location.href='ro.asp?roid=<%=request("roid")%>&showrec=y'
		}else{
			alert(x)
		}
	}
}

function addRecRepair(){
	document.getElementById("recrepairs").style.display = "none"
	document.getElementById("addrecrepairs").style.display = "block"
}
function cancelAddRec(){
	document.getElementById("addrecrepairs").style.display = "none"
	document.getElementById("popuphider").style.display = "none"
}
function addRecRep(roid){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	qs = "&desc=" + document.getElementById("desc").value + "&totalrec=" + document.getElementById("recamt").value
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	var url="addrecrep.asp?roid="+roid+qs
	//alert(url)
	xmlHttp.onreadystatechange=addRecChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}
function addRecChanged(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		x = xmlHttp.responseText
		if(x == "success"){
			document.body.onunload = ""
			location.href='ro.asp?roid=<%=request("roid")%>&showrec=y'
		}else{
			alert(x)
		}
	}

}
function editRecRep(){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	qs = "?id="+document.getElementById("editrecid").value+"&desc=" + document.getElementById("editdesc").innerHTML + "&totalrec=" + document.getElementById("editrecamt").value
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	var url="editrecrep.asp"+qs
	//alert(url)
	xmlHttp.onreadystatechange=editRecChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}
function editRecChanged(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		x = xmlHttp.responseText
		if(x == "success"){
			document.body.onunload = ""
			location.href='ro.asp?roid=<%=request("roid")%>&showrec=y'
		}else{
			alert(x)
		}
	}

}
function editRec(id,d,a){
	document.getElementById("recrepairs").style.display = "none"
	document.getElementById("editrecrepairs").style.display = "block"
	document.getElementById("editrecamt").value = a
	document.getElementById("editrecid").value = id
	document.getElementById("editdesc").value = d

}
function cancelEditRec(){
	document.getElementById("editrecrepairs").style.display = "none"
	document.getElementById("popuphider").style.display = "none"

}
function declineItem(comid){
	c = confirm("This will decline the repairs for this issue AND move all parts and labor to the Recommended Repairs.  Are you sure?")
	if (c){
		document.body.onunload='';
		document.theform.action='savero.asp?comid='+comid+'&comstat=Declined';
		saveAll('ro.asp?a=n','','','','','')
	}

}

function changePos(comid,currpos,direction,roid){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	//alert(alldataurl)
	var url="complaintreorder.asp?roid=<%=request("roid")%>&id="+comid+"&currpos="+currpos+"&direction="+direction
	r=Math.floor(Math.random()*111111)
	url += "&r=" + r
	xmlHttp.onreadystatechange=comStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}

function comStateChanged(){

	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		if(xmlHttp.responseText != "error"){
			//alert(xmlHttp.responseText)
			document.body.onunload='';
			saveAll('ro.asp?a=n','','','','','')

		}else{
			alert(xmlHttp.responseText)
		}
	}
	
}

function timeClock(laborid,comid,tech){
	//alert(document.getElementById("clock"+laborid).innerHTML.indexOf("clock.png"))
	if (document.getElementById("clock"+laborid).innerHTML.indexOf("clock.png") >= 27){
		// clocking in
		document.getElementById("clock"+laborid).innerHTML = '<img alt="" height="17" src="newimages/runningman.gif" width="17">Clock Out'
		dar = getLocalDate()
		dar = dar.split("*")
		setCookie(dar[0],dar[1],"1")

	}else{
		// clocking out
		document.getElementById("clock"+laborid).innerHTML = '<img alt="" src="newimages/clock.png">Clock In'
		dar = getLocalDate()
		dar = dar.split("*")
		getTechForTime(tech,comid,laborid,getCookie("starttime"),dar[1])
		//postLaborTimeClock(getCookie("starttime"),dar[1],<%=request("roid")%>,comid,laborid,tech)
		
		//now call via ajax to post the start and end date/time to the database
	}
	//alert(getLocalDate())
}

function getLocalDate(){
	var currentTime = new Date()
	var month = currentTime.getMonth() + 1
	var day = currentTime.getDate()
	var year = currentTime.getFullYear()
	
	var currentTime = new Date()
	var hours = currentTime.getHours()
	var minutes = currentTime.getMinutes()
	var secs = currentTime.getSeconds()

	return "starttime*"+month+"/"+day+"/"+year+" "+hours+":"+minutes+":"+secs

}

function getCookie(c_name)
 {
 var i,x,y,ARRcookies=document.cookie.split(";");
 for (i=0;i<ARRcookies.length;i++)
 {
   x=ARRcookies[i].substr(0,ARRcookies[i].indexOf("="));
   y=ARRcookies[i].substr(ARRcookies[i].indexOf("=")+1);
   x=x.replace(/^\s+|\s+$/g,"");
   if (x==c_name)
     {
     return unescape(y);
     }
   }
 }

function setCookie(c_name,value,exdays){
	 var exdate=new Date();
	 exdate.setDate(exdate.getDate() + exdays);
	 var c_value=escape(value) + ((exdays==null) ? "" : "; expires="+exdate.toUTCString());
	 document.cookie=c_name + "=" + c_value;
 }

function getTechForTime(tech,comid,laborid,startdate,enddate){

	//techfortime
	document.getElementById("popuphider").style.display = "block"
	document.getElementById("techfortime").style.display = "block"
	document.techfortimeform.comid.value=comid
	document.techfortimeform.laborid.value=laborid
	document.techfortimeform.startdate.value = startdate
	document.techfortimeform.enddate.value = enddate
	
	// set the selected technician
	var textToFind = tech; 
	 
	var dd = document.techfortimeform.techformtimetech; 
	for (var i = 0; i < dd.options.length; i++) { 
	    if (dd.options[i].text == textToFind) { 
	        dd.selectedIndex = i; 
	        break; 
	    } 
	} 

}

function postLaborTimeClock(){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="labortimeclockpost.asp?startdate="+document.techfortimeform.startdate.value+"&enddate="+document.techfortimeform.enddate.value+"&roid=<%=request("roid")%>&comid="+document.techfortimeform.comid.value+"&laborid="+document.techfortimeform.laborid.value+"&tech="+document.techfortimeform.techformtimetech.value
	//alert(url)
	xmlHttp.onreadystatechange=laborTimeStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function laborTimeStateChanged(){

	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		document.getElementById("popuphider").style.display = "none"
		document.getElementById("techfortime").style.display = "none"
		document.body.onunload = ""
		saveAll("ro.asp?a=n")
	}
}

function showClockList(laborid,tobj){

	document.getElementById("hider").style.display = "block"
	//document.getElementById("clocklist1-"+laborid).style.display = "block"
	document.getElementById("clocklist2-"+laborid).style.position = "absolute"
	curtopleft = findMouse(tobj).toString()
	tar = curtopleft.split(",")
	
	document.getElementById("clocklist2-"+laborid).style.top = tar[1]
	document.getElementById("clocklist2-"+laborid).style.left = tar[0]
	document.getElementById("clocklist2-"+laborid).style.display = "block"

}

function closeClockList(laborid){

	document.getElementById("hider").style.display = "none"
	//document.getElementById("clocklist1-"+laborid).style.display = "none"
	document.getElementById("clocklist2-"+laborid).style.display = "none"

}

function findMouse(obj) {

 
  var curleft = curtop = 0, scr = obj, fixed = false;
  while ((scr = scr.parentNode) && scr != document.body) {
    curleft -= scr.scrollLeft || 0;
    curtop -= scr.scrollTop || 0;
    if (getStyle(scr, "position") == "fixed") fixed = true;
  }
  if (fixed && !window.opera) {
    var scrDist = scrollDist();
    curleft += scrDist[0];
    curtop += scrDist[1];
  }
  do {
    curleft += obj.offsetLeft;
    curtop += obj.offsetTop;
  } while (obj = obj.offsetParent);
  return [curleft, curtop];

}

function showParts(v){

	window.open(v,"po","toolbar=no,menubar=no")

}

function deleteVehIssue(id){

	c = confirm("This will delete this issue and all parts and labor associated with it.  Are you sure?")
	if(c){
		saveAll('deletecomplaintfromro.asp?complaintid='+id+"&",'','','','','')
	}
	
}

function setReminder(){

	document.getElementById("popuphider").style.display = "block"
	document.getElementById("reminders").style.display = "block"

}

function closeReminder(){
	document.getElementById("popuphider").style.display = "none"
	document.getElementById("reminders").style.display = "none"

}

function addReminder(){
	
	document.getElementById("reminders").style.display = "none"
	document.getElementById("popuphider").style.display = "block"
	document.getElementById("addreminder").style.display = "block"
}

function closeAddReminder(){

	document.getElementById("popuphider").style.display = "none"
	document.getElementById("addreminder").style.display = "none"


}

function saveReminder(){
	

	if (document.remform.reminder.value.length == 0){
		alert("Please enter the reason for the reminder")
		return
	}
	
	if (document.remform.reminderdate.value.length == 0){
		alert("Please enter the date for the reminder")
		return
	}
	document.theform.action=""
	document.remform.submit()

}

function editDate(){

	document.getElementById("datebox").style.display = "block"

}
function printDiv(divName) {
	document.getElementById("b1").style.display = 'none'
	document.getElementById("b2").style.display = 'none'
	document.getElementById("b3").style.display = 'none'
     var printContents = document.getElementById(divName).innerHTML;
     var originalContents = document.body.innerHTML;

     document.body.innerHTML = printContents;

     window.print();

     document.body.innerHTML = originalContents;
     
	document.getElementById("b1").style.display = 'block'
	document.getElementById("b2").style.display = 'block'
	document.getElementById("b3").style.display = 'block'
	document.getElementById("b1").style.right = '328px'

}
function showPOs(showhidebutton){
	r = Math.random()
	document.getElementById("pos").src = "roshowpo.asp?roid=<%=request("roid")%>&r=" + r
	document.getElementById("pos").style.display = "block"
	document.getElementById("popuphider").style.display = "block"
	if(showhidebutton == "hide"){
		document.getElementById("backbutton").style.display = 'none'
	}
}

</script>
</head>
<%
if request("showrec") = "y" then srr = "showRecRepairs()"
%>

<body link="#800000" vlink="#800000" alink="#800000" onload="ttlFees();closeTimer();<%=srr%>"  onunload="javascript:saveAll('wip.asp?a=n','','','','','')" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<input id="backbutton" style="font-size:18px;z-index:999;height:5%;color:red;font-weight:bold;display:none;position:absolute;top:5px;left:5%;" onclick="showPOs('hide')" type="button" onclick="" value="Back To PO Managment">
<iframe id="pos" style="width:90%;position:absolute;left:5%;top:5%;height:90%;background-color:white;display:none;z-index:1000"></iframe>

<iframe src="" style="display:none" id="inspection"></iframe>
<input name="dfield" type="hidden" ><input name="caldatefield" type="hidden" id="caldatefield">
<div style="z-index:1000;display:none;width:60%;position:absolute;left:20%;top:100px;height:500px;background-color:white;padding:20px;" id="reminders">

	<strong>Service Reminders - <span onclick="addReminder();document.remform.reminder.focus()" style="color:#0258BD;cursor:pointer;padding:10px;border:medium black outset;background-color:#3366CC;color:white;border:thin navy outset">Add Reminder</span>
	<input onclick="closeReminder()" style="font-size:xx-small;width:40px;position:absolute;top:15px;right:15px;cursor:pointer;height:30px" name="Button1" type="button" value="Close" >
	</strong><br><br>
	<table style="width: 100%" cellpadding="4" cellspacing="0" >
		<tr>
			<td class="style10" style="height: 46; width: 75%;"><strong>Service&nbsp;</strong></td>
			<td class="style10" style="height: 46; width: 30%;"><strong>Date&nbsp;</strong></td>
		</tr>
		<%
		stmt = "select * from reminders where shopid = '" & request.cookies("shopid") & "' and customerid = " & rs("customerid") & " and vehid = " & rs("vehid") _
		& " order by reminderdate asc"
		'response.write stmt
		set remrs = con.execute(stmt)
		if not remrs.eof then
			do until remrs.eof
		%>
		<tr>
			<td style="width: 75%; height: 24px"><%=remrs("reminderservice")%>&nbsp;</td>
			<td style="height: 24px"><%=remrs("reminderdate")%>&nbsp;</td>
		</tr>
		<%
				remrs.movenext
			loop
		else
		%>
		<tr>
			<td colspan="2">No service reminders set&nbsp;</td>
		</tr>
		<%
		end if
		%>
	</table>


</div>

<div style="z-index:1000;display:none;width:60%;position:absolute;left:20%;top:100px;height:500px;background-color:white;padding:20px;border:medium black outset" id="addreminder">
	<form action="addreminder.asp" method="post" name="remform">
	<input name="customerid" type="hidden" value="<%=rs("customerid")%>" ><input name="vehid" type="hidden" value="<%=rs("vehid")%>">
	<input name="roid" type="hidden" value="<%=request("roid")%>">
	<strong>Enter the Service Reminder and Date<br></strong>&nbsp;<table style="width: 100%" cellpadding="4" cellspacing="0">
		<tr>
			<td class="style10" style="height: 44px">Service Req'd&nbsp;</td>
			<td class="style10" style="height: 44px">Date Due&nbsp;</td>
		</tr>
		<tr>
			<td><input name="reminder" type="text" style="width: 437px;text-transform:capitalize" >&nbsp;</td>
			<td>&nbsp;<input onfocus="show_cal(this,'demo1');" id="demo1" name="reminderdate" style="width: 120px" type="text"> <img src="javascripts/images2/cal.gif" onclick="show_cal(this,'demo1');" style="cursor:pointer"></td>
		</tr>
	</table>
	<input onclick="saveReminder()" name="Button1" type="button" value="Add Reminder" > <input onclick="closeAddReminder()" name="Button1" type="button" value="Cancel" >
	</form>
</div>

<table id="toptable" border=2 cellpadding=0 cellspacing=0 width=100% bordercolor="#0066CC" style="border-collapse: collapse"><tr><td>
<form name="theform" method="post" action="savero.asp">

<div id="addrecrepairs">
	
	<table style="width: 100%">
		<tr>
			<td valign="top"><strong>Amount of Recommended Repair</strong>&nbsp;</td>
			<td valign="top"><input name="recamt" type="text" id="recamt" >&nbsp;</td>
		</tr>
		<tr>
			<td valign="top"><strong>Description of Recommended Repairs</strong>&nbsp;</td>
			<td valign="top"><textarea id="desc" style="height:300px;width:400px" name="recommend" cols="20" rows="2"></textarea></td>
		</tr>
		<tr>
			<td><input onclick="addRecRep('<%=request("roid")%>')" name="Button1" type="button" value="Add Recommendation" > <input onclick="cancelAddRec()" name="Button1" type="button" value="Cancel" ></td>
		</tr>
		
	</table>

</div>
<div id="editrecrepairs">
	<input name="editrecid" id="editrecid" type="hidden" />
	<table style="width: 100%">
		<tr>
			<td valign="top"><strong>Amount of Recommended Repair</strong>&nbsp;</td>
			<td valign="top"><input name="recamt" type="text" id="editrecamt" >&nbsp;</td>
		</tr>
		<tr>
			<td valign="top"><strong>Description of Recommended Repairs</strong>&nbsp;</td>
			<td valign="top"><textarea id="editdesc" style="height:300px;width:400px" name="editdesc" cols="20" rows="2"></textarea></td>
		</tr>
		<tr>
			<td><input onclick="editRecRep()" name="Button1" type="button" value="Save Changes" > <input onclick="cancelEditRec()" name="Button1" type="button" value="Cancel" ></td>
		</tr>
		
	</table>

</div>

<div id="recrepairs">
<input id="b1" style="position:absolute;top:5px;right:328px;cursor:pointer" onmousedown="this.style.right='318px';" name="Button1" type="button" onclick="printDiv('recrepairs')" value="Print" >
<input id="b2" style="position:absolute;top:5px;right:20px;cursor:pointer" onmousedown="this.style.right='3px';" name="Button1" type="button" onclick="closeRecRepairs()" value="Close" >
<input id="b3" style="position:absolute;top:5px;right:75px;cursor:pointer" onmousedown="this.style.right='58px';" name="Button1" type="button" onclick="addRecRepair()" value="Add Recommended Repair" >
Recommended Repairs:

	<table cellpadding="2" cellpadding="0" style="width: 100%">
		<tr>
			<td style="background-color:maroon;color:white;font-weight:bold">Recommendation&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold;text-align:right">Repair Cost&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold;text-align:right">Restore/Delete&nbsp;</td>
		</tr>
		<%
		stmt = "select * from recommend where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid") 
		set recrs = con.execute(stmt)
		if not recrs.eof then
			do until recrs.eof
				
				' test for parts or labor
				showrestore = "no"
				stmt = "select * from recommendparts where recid = " & recrs("id")
				set trec = con.execute(stmt)
				if not trec.eof then
					showrestore = "yes"
				end if
				if showrestore = "no" then
					stmt = "select * from recommendlabor where recid = " & recrs("id")
					set trec = con.execute(stmt)
					if not trec.eof then
						showrestore = "yes"
					end if
				end if
		%>
		<tr>
			<td id="desc" style="font-weight:bold"><%=recrs("desc")%></td>
			<td id="amount" style="font-weight:bold;text-align:right"><%=formatnumber(recrs("totalrec"),2)%></td>
			<td style="font-weight:bold;text-align:right">
			<input onclick="editRec('<%=recrs("id")%>','<%=recrs("desc")%>','<%=formatnumber(recrs("totalrec"),2)%>')" name="Button1" type="button" value="Edit">
			<%
			if showrestore = "yes" then
			%>
			<input onclick="document.body.onunload='';document.theform.action='savero.asp?comid=<%=recrs("comid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" name="Button1" type="button" value="Restore" > 
			<%
			end if
			%>
			<input onclick="deleteRec('<%=recrs("id")%>')" name="Button1" type="button" value="Delete" >
			</td>
		</tr>
		<%
				recrs.movenext
			loop
		else
		%>
		<tr>
			<td style="font-weight:bold">No Recommendations</td>
			<td style="font-weight:bold"></td>
			<td style="font-weight:bold"></td>
		</tr>
		<%
		end if
		%>
	</table>
<div style="text-align:left;height:400px">
</div>

</div>
<div id="customvehiclefields">
<%
if len(rs("customvehicle1label")) > 0 then
	response.write rs("customvehicle1label") & ":<br>" & rs("customvehicle1") & "<br><br><br>"
end if
if len(rs("customvehicle2label")) > 0 then
response.write rs("customvehicle2label") & ":<br>" & rs("customvehicle2") & "<br><br><br>"
end if
if len(rs("customvehicle3label")) > 0 then
response.write rs("customvehicle3label") & ":<br>" & rs("customvehicle3") & "<br><br><br>"
end if
if len(rs("customvehicle4label")) > 0 then
response.write rs("customvehicle4label") & ":<br>" & rs("customvehicle4") & "<br><br><br>"
end if
%>
<input name="Button1" type="button" onclick="hideVehicleFields()" value="Close" />
</div>
<input type="hidden" name="x" value="" >
<input type="hidden" name="oldstatus" value="<%=rs("status")%>" >
 <input type="hidden" name="dir" value><input type="hidden" name="itype" value><input type="hidden" name="pid" value><input type="hidden" name="sid" value><input type="hidden" name="lid" value>
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
   <tr>
    <td width=" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
       <tr>
        <td valign="top" width="100%">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
           <tr>
            <td valign="top" height="22" >
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
               <td width="100%" align="center" class="style10" ><b><i>
               <font color="#FFFF00" size="5">RO#&nbsp;<%=roid%></font></i></b><img style="position:absolute;right:5px;top:0px;" alt="" height="37" src="newimages/shopbosslogo-ro.png" width="157"></td>
              </tr>
              <tr>
               <td width="100%" align="left" bgcolor="#0066CC">

             
               <table style="width: 100%">
				   <tr>
					   <td onclick="document.theform.Status.focus();document.body.onunload='';saveAll('ro.asp?a=n','','','','','')"  class="menustyle">
					   <img align="middle" alt="" height="30" src="newimages/icon-calculator.jpg" width="31"><br>Calculate</td>
					   <td onclick="saveAll('pdfinvoices/printpdfro.asp?roid=<%=request("roid")%>&','','','','');" valign="middle" class="menustyle">
					   <img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br>
					   Print RO</td>
					   <td onclick="showHistory()" class="menustyle">
					   <img alt="" height="30" src="newimages/historyx30.png" width="30"><br>
					   History</td>
					   <td onclick="showInspection()" class="menustyle">
					   <img alt="" height="29" src="newimages/inspection.gif" width="40"><br>
					   Inspection</td>
					   <td style="z-index:999" onclick="showPOs('')" class="menustyle">
					   <img alt="" height="30" src="newimages/icon_salesorders.gif" width="31"><br>
					   					   PO Number</td>

					   <td onclick="uploadFiles()" class="menustyle">
					   <img alt="" height="30" src="newimages/upload.png" width="39"><br>Upload/View Pics</td>
					   <td onclick="setReminder()" class="menustyle">
					   <img alt="" height="30" src="newimages/reminder.png" width="30"><br>
					   Set/View Reminders</td>
					   <td onclick="showUpdateWindow('y')" class="menustyle">
					   <img alt="" height="30" src="newimages/alert.png" width="30"><br>
					   Send Update</td>
					   <td onclick="document.body.onunload='';location.href='complaintaddupdate.asp?roid=<%=request("roid")%>'" class="menustyle">
					   <img alt="" height="30" src="newimages/vintage-car-icon-ro.png" width="45">&nbsp;<br>
					   &nbsp;Vehicle Issues</td>
					   <td onclick="showRecRepairs()" class="menustyle">
					   <img alt="" height="30" src="newimages/wrench.png" width="45">&nbsp;<br>
					   &nbsp;Recommended Repairs</td>
					   <td onclick="document.theform.Status.focus();location.href='ro.asp?roid=<%=request("roid")%>';saveAll('wip.asp?a=n','','','','','')" class="menustyle">
					   <img alt="" height="30" src="newimages/save.png" width="30">&nbsp;<br>
					   Save/Done</td>
				   </tr>
			   </table>

             
               </td>
              </tr>

             </table>
             </td>
           </tr>
          </table>
         <center>
         <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
           <td width="34%" valign="top" bgcolor="#ffffff" height="170">
           <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" height="100%" class="style16">
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              RO/Tag #</td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15">
			  <span id='tagnumber'><%=ucase(rs("tagnumber"))%></span>&nbsp;&nbsp;
			  <%
			  if len(rs("tagnumber")) > 0 then
			  %>
			  <span id="deltagnumber" onclick="document.theform.dbtagnumber.value='';document.getElementById('tagnumber').style.display='none';document.getElementById('deltagnumber').style.display='none'" style="font-weight:bold;font-size:x-small;color:blue;cursor:pointer">(Delete Tag #)</span></td>
              <%
              end if
              %>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14" style="height: 21px">
              <font size="2" face="Arial">Customer:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15" style="height: 21px"><font size="2" face="arial"><%=rs("Customer")%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Address:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=rs("CustomerAddress")%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">City,State,Zip:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=rs("CustomerCSZ")%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Home:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=homephone%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Work:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=workphone%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Cell:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=cellphone%></font>&nbsp;</td>
             </tr>
             <tr>
              <td style="font-family:Arial;font-size:x-small;font-weight:normal" width="10%" align="right" valign="bottom" class="style1">
              Bus. Contact</td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><%=rs("contact")%>&nbsp;</td>
             </tr>
            </table></td>
           <td width="33%" valign="top" bgcolor="#FFFFFF" height="170">
           <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" height="100%" class="style16">
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Vehicle:</font></td>
              <td style="overflow:hidden" width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=left(rs("VehInfo"),27)%></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14" style="height: 23px">
              <font size="2" face="Arial">License:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 23px"><font size="2" face="arial"><%=rs("VehLicNum")%>&nbsp;&nbsp;&nbsp;Fleet #<%=rs("fleetno")%></font></td>
             </tr>
             <tr>
              <td onclick="saveAll('pdfinvoices/1122/printpdfronew.asp?roid=<%=request("roid")%>&','','','','');"" width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">VIN:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("Vin")%></font>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial"><%=milesinlabel%>:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15">
               <font size="2" face="arial">
               <input onkeyup="chgDB(this.name);" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" id="MilesIn" name="VehicleMiles" value="<%=rs("VehicleMiles")%>" style="padding: 1 4; border: 1px solid #BBD7F2; height:18px; font-family:Arial; font-size:10pt; text-transform:uppercase; " size="9"></b></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14" style="height: 24px">
              <font size="2" face="Arial"><%=milesoutlabel%>:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 24px">
               <font size="2" face="arial"><b>
               <input onkeyup="chgDB(this.name);" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" id="MilesOut" name="MilesOut" value="<%=rs("MilesOut")%>" style="border:1px solid #BBD7F2; height:18px; font-family:Arial; font-size:10pt; font-variant:small-caps; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1" size="9"></b></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Engine/Drive:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("VehEngine")%>&nbsp;-&nbsp;<%=rs("Cyl")%>Cyl&nbsp;-&nbsp;<%=rs("VehTrans")%>&nbsp;-&nbsp;<%=rs("DriveType")%></font></td>
             </tr>
             <tr>
              <td colspan="2" style="cursor:pointer;background-color:white" width="20%" valign="top" class="style24">
			  <input onclick="showVehicleFields()" name="Button1" type="button" value="Custom Vehicle Fields" >
			  </td>
             </tr>
            </table></td>
           <td width="33%" valign="top" height="170">
           <table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" height="100%">
             <tr>
              <td ondblclick="editDate()" width="10%" align="right" valign="top" height="44" style="height: 22px" class="style14">
			  RO Date</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="44" style="height: 22px" class="style15">
			  <%=rs("datein")%>
			  <div id="datebox" style="display:none;font-size:small">
			  In:<input onfocus="show_cal(this,'DateIn');document.getElementById('caldatefield').value='DateIn'" id="DateIn" size="10" onblur="" name="DateIn" value="<%=rs("datein")%>" type="text" > Status:<input id="StatusDate" onfocus="show_cal(this,'StatusDate');document.getElementById('caldatefield').value='StatusDate'" size="10" onblur="" name="StatusDate" value="<%=rs("statusdate")%>" type="text" >
			  </div>
			  </td>
             </tr>
             <tr>
              <td ondblclick="alert(document.theform.dbStatusDate.value+':'+document.theform.dbDateIn.value)" width="10%" align="right" valign="top" height="44" style="height: 22px" class="style14">
			  Status</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="44" style="height: 22px" class="style15">
			  <select style="height:18px;" onchange="setStatusDate();chgDB(this.name);closeRO()" size="1" name="Status" >
              <option selected value="<%=ucase(rs("Status"))%>">
              <%
              if ucase(rs("status")) = "FINAL" then
              	response.write rs("status")
              else
	            response.write mid(ucase(rs("Status")),2,len(rs("status"))-1)
	          end if
              %>
              </option>
              <%
              set statusrs = con.execute("select * from statuses where `status` != '" & rs("status") & "' order by `order`")
              do until statusrs.eof
              	s = statusrs("status")
              	if isnumeric(left(s,1)) then
              		displays = right(s,len(s)-1)
              	else
              		displays = s
              	end if
              %>
              <option value="<%=s%>"><%=displays%></option>
              <%
              	statusrs.movenext
              loop
              set statusrs = nothing
              %>
             </select></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" height="31" class="style14"><font size="2" face="arial">Comments:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="31" class="style15"><font size="2" face="arial">
              <textarea onfocus="javascript:this.style.backgroundColor='#FFFFCC'" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white';" name="Comments" cols="37" style="padding: 1 4; text-transform: uppercase; font-size: 8pt; font-family: Arial; height: 39px;" tabindex="21"><%=rs("Comments")%></textarea></font></td>
             </tr>
             <tr>
              <td ondblclick="saveAll('pdfinvoices/printpdfro-new.asp?roid=<%=request("roid")%>&','','','','')" width="10%" align="right" valign="top" class="style14" style="height: 22px"><font size="2" face="arial">Writer:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 22px">
               <select onchange="javascript:chgDB(this.name);" id="Writer" size="1" name="Writer">
                <option selected value="<%=rs("Writer")%>"><%=rs("Writer")%></option>
                <%
				estmt = "Select EmployeeFirst, EmployeeLast from employees where shopid = '" & request.cookies("shopid") & "' and Active = 'Yes'"
				set ers = con.execute(estmt)
				if not ers.eof then
				ers.movefirst
				while not ers.eof
					ematch = ers("EmployeeFirst") & " " & ers("employeelast")
					if rs("writer") = ematch then
						s = " selected='selected' "
					else
						s = ""
					end if
				response.write "<option " & s & " value='" & ers("EmployeeFirst") & " " & ers("EmployeeLast") & "'>" & ers("EmployeeFirst") & " " & ers("EmployeeLast") & "</option>"
				ers.movenext
				wend
				end if
				set ers = nothing
                %>
               </select></td>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14"><font size="2" face="arial">Source:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15"><font size="2" face="arial">
               </font><select size="1" onchange="javascript:chgDB(this.name);" id="Source" name="Source" >

                <%
				srcstmt = "select Source from source where shopid = '" & request.cookies("shopid") & "'"
				set srcrs = con.execute(srcstmt)
				srcrs.movefirst
				while not srcrs.eof
					if srcrs("source") = rs("source") then
						s = " selected='selected' "
					else
						s = ""
					end if
					response.write "<option " & s & " value='" & srcrs("Source") & "'>" & srcrs("Source") & "</option>"
					srcrs.movenext
				wend
				set srcrs = nothing
                %>
               </select></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14">
               <p align="right" class="style13"><font size="2" face="arial">Type:</font></td>


              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15"><select size="1" onchange="javascript:chgDB(this.name);" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white'" id="ROType0" name="ROType" size="20">
                <option selected value="<%=rs("ROType")%>"><%=rs("ROType")%></option>
                <%
				srcstmt = "select ROType from rotype where shopid = '" & request.cookies("shopid") & "'"
				set srcrs = con.execute(srcstmt)
				srcrs.movefirst
				while not srcrs.eof
					response.write "<option value='" & srcrs("ROType") & "'>" & srcrs("ROType") & "</option>"
					srcrs.movenext
				wend
				set srcrs = nothing
                %>
               </select></td>
             </tr>
            </table></td>
          </tr>
          <tr>
           <td width="67%" valign="top" bgcolor="#FFFFFF" colspan="2">
		   <div id="issues" style="max-height:450px;overflow:auto">
           <!-- #include file=1122complaintlayout.asp -->
		   </div>

           </td>
           <td width="33%" valign="top">
           <table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse; height:30px;" bordercolor="#111111">
             <tr>
              <td style="cursor:pointer" id="revcell" Width="25%" align="center" onclick="revChg()" class="style9">
			  <img alt="" height="30" src="newimages/desktop.png" width="30"><br>Revisions</td>
              <td style="cursor:pointer" id="feecell" width="25%" align="center" onclick="feesChg()" class="style9">
			  <img alt="" height="30" src="newimages/discount_icon.png" width="28"><br>Fees/Disc</td>
              <td style="cursor:pointer" id="warrcell" width="25%" align="center" onclick="warrChg()" class="style9">
			  <img alt="" height="30" src="newimages/warricon2.png" width="32"><br>Warranty</td>
              <td style="cursor:pointer" id="pmtscell" width="25%" align="center"onclick="pmtsChg()" class="style9">
			  <img alt="" height="30" src="newimages/payment.png" width="30"><br>Payments</td>
             </tr>
            </table><table cellspacing="0" cellpadding="0" width="100%" border="0">
             <tr>
              <td id="est" style="display:block">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td width="50%" colspan="2" class="style17">
                  <p align="right" class="style13">Original Estimate:</p>
                 </td>
                 <td width="50%" colspan="2" bgcolor="#FFFFFF" class="style15"><input style="height: 18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="OrigRO" size="9" value="<%=formatnumber(rs("OrigRO"),2)%>"></td>
                </tr>
                <tr>
                 <td width="50%" colspan="2" class="style17">
                  <p align="center" class="style13"><u>Revision 1</u></p>
                 </td>
                 <td width="50%" colspan="2" class="style17">
                  <p align="center" class="style13"><u>Revision 2</u></p>
                 </td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">Amt:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height:18; font-variant: small-caps; font-family: Arial; font-size: 8pt; text-align: Right; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1Amt" size="9" value="<%=formatnumber(rs("Rev1Amt"),2)%>" tabindex="10"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">Amt:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Amt" size="9" value="<%=formatnumber(rs("Rev2Amt"),2)%>" style="font-family: Arial; font-variant: small-caps; font-size: 8pt; text-align: Right; height:18; " tabindex="15"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onkeyup="chgDB(this.name);" onblur="chgDB(this.name)" type="text" name="Rev1Date" size="9" value="<%=rs("Rev1Date")%>" tabindex="11"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Date" size="9" value="<%=rs("Rev2Date")%>" tabindex="16"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onkeyup="chgDB(this.name);" onblur="chgDB(this.name)" type="text" name="Rev1Phone" size="9" value="<%=rs("Rev1Phone")%>" tabindex="12"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Phone" size="9" value="<%=rs("Rev2Phone")%>" tabindex="17"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1Time" size="9" value="<%=rs("Rev1Time")%>" tabindex="13"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Time" size="9" value="<%=rs("Rev2Time")%>" tabindex="18"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1By" size="9" value="<%=rs("Rev1By")%>" tabindex="14"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2By" size="9" value="<%=rs("Rev2By")%>" tabindex="19"></font></td>
                </tr>
               </table></td>
             </tr>
             <tr>
              <td id="fees" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td ondblclick="calcPercents()" width="50%" align="right" class="style14"><font size="2">Haz. Waste
                  Fee:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
                  $ <input style="height:18; font-family: Arial; font-variant: small-caps; text-align:right; font-size: 8pt; width:60px; " onblur="setZero(this.name,this.value);chgDB(this.name);ttlFees();this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="HazardousWaste" size="9" value="<%=formatdollar(rs("HazardousWaste"))%>">
                 </td>
                </tr>
                 <%
                 'get userfees from company
                 
                 if formatnumber(rs("userfee1")) >= 0 then
                 	if len(rs("userfee1label")) > 0 then
                		userfee1desc = rs("userfee1label") & ":"
                	end if
                 	userfee1amount = formatnumber(rs("userfee1"),2)
                 	userfee1percent = rs("userfee1percent")
                 	if rs("userfee1type") = "$" then
                 		userfee1display = "$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='setZero(this.name,this.value);chgDB(this.name);ttlFees();' type='text' id='UserFee1' name='UserFee1' value='" & userfee1amount & "'>"
                 	elseif rs("userfee1type") = "%" then
                  		userfee1display = "<a href='javascript:void(0)' onclick='clearFee(""1"")'>Clear</a>&nbsp;(" & userfee1percent & "%)&nbsp;&nbsp;&nbsp;&nbsp;$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='' type='text' id='UserFee1' name='UserFee1' value='" & userfee1amount & "'>"
                 	end if
                 else
                 	userfee1desc = ""
                 	userfee1display = ""
                 end if
                 
                 if formatnumber(rs("userfee2")) >= 0 then
                 	if len(rs("userfee2label")) > 0 then
                		userfee2desc = rs("userfee2label") & ":"
                	end if

                 	userfee2amount = formatnumber(rs("userfee2"),2)
                 	userfee2percent = rs("userfee2percent")
                 	if rs("userfee2type") = "$" then
                 		userfee2display = "$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='setZero(this.name,this.value);chgDB(this.name);ttlFees();' type='text' id='UserFee2' name='UserFee2' value='" & userfee2amount & "'>"
                 	elseif rs("userfee2type") = "%" then
                  		userfee2display = "<a href='javascript:void(0)' onclick='clearFee(""2"")'>Clear</a>&nbsp;(" & userfee2percent & "%)&nbsp;&nbsp;&nbsp;&nbsp;$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='' type='text' id='UserFee2' name='UserFee2' value='" & userfee2amount & "'>"
                 	end if
                 else
                 	userfee2desc = ""
                 	userfee2display = ""
                 end if

                 if formatnumber(rs("userfee3")) >= 0 then
                 	if len(rs("userfee3label")) > 0 then
                		userfee3desc = rs("userfee3label") & ":"
                	end if

                 	userfee3amount = formatnumber(rs("userfee3"),2)
                 	userfee3percent = rs("userfee3percent")
                 	if rs("userfee3type") = "$" then
                 		userfee3display = "$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='setZero(this.name,this.value);chgDB(this.name);ttlFees();' type='text' id='UserFee3' name='UserFee3' value='" & userfee3amount & "'>"
                 	elseif rs("userfee3type") = "%" then
                  		userfee3display = "<a href='javascript:void(0)' onclick='clearFee(""3"")'>Clear</a>&nbsp;(" & userfee3percent & "%)&nbsp;&nbsp;&nbsp;&nbsp;$ <input size='13' style=' width:60px;font-size:10px;font-variant:small-caps;text-align:right;height:18px' onblur='' type='text' id='UserFee3' name='UserFee3' value='" & userfee3amount & "'>"
                 	end if
                 else
                 	userfee3desc = ""
                 	userfee3display = ""
                 end if

                 %>

                <tr>
                 <td width="50%" align="right" class="style14"><%=userfee1desc%>&nbsp;</td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=userfee1display%></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><%=userfee2desc%>&nbsp;</td>
                   <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=userfee2display%></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><%=userfee3desc%>&nbsp;</td>

                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=userfee3display%></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">Storage
                  Fee:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
				 $
				 <input style="height:18px; font-family: Arial; font-variant: small-caps; font-size: 8pt;text-align: right; width:60px;" onblur="setZero(this.name,this.value);chgDB(this.name);ttlFees();this.style.backgroundColor='white'" type="text" name="storagefee" size="9" value='<%=formatdollar(rs("storagefee"))%>'></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">Parts Tax Rate:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><input style="height:18px; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="TaxRate" size="6" value="<%=rs("TaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14">Labor Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
                 <input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="LaborTaxRate" size="6" value="<%=rs("LaborTaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14">Sublet Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
                 <input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="SubletTaxRate" size="6" value="<%=rs("SubletTaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">Discount
                  Amt:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
				 $ <input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; width:60px; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white';adjustSalesTax(this.value);" type="text" onkeyup="chgDB(this.name);" style="text-align:right" name="DiscountAmt" size="9" value="<%=formatdollar(rs("DiscountAmt"))%>"></td>
                </tr>
                </table></td>
             </tr>
             <tr>
              <td id="warr" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td  width="50%" align="right" class="style17"><font size="2">Warranty
                  Months:&nbsp;</font></td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" onkeyup="chgDB(this.name);" style="text-align:right" name="WarrMos" size="7" value="<%=rs("WarrMos")%>"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17"><font size="2">Warranty
                  Miles:&nbsp;</font></td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" style="text-align:right" name="WarrMiles" size="7" value="<%=rs("WarrMiles")%>"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="visibility:hidden; height:18" type="text" name="T26" size="7"></td>
                </tr>
                <tr>
                 <td width="50%" class="style25">
                 <input onclick="showDisclosure()" name="Button2" type="button" value="Disclosures">
				 </td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="visibility:hidden; height:18" type="text" name="T28" size="7"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="visibility:hidden; height:18" type="text" name="T27" size="7"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="visibility:hidden; height:18" type="text" name="T29" size="7"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18">&nbsp;</td>
                </tr>
               </table></td>
             </tr>
             <tr>
              <td id="pmts" style="display:none" class="style15">
              <table width="100%" border="1" cellspacing="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
                 <td class="style21" style="width: 33%">
				 Paid with</td>
                 <td width="33%" class="style21" style="width: 33%">
				 Date</td>
                 <td width="33%" align="right" class="style14" style="width: 33%">
				 Amount</td>
				 </tr>
				 <%
				 stmt = "select * from accountpayments where shopid = '" & request.cookies("shopid") _
				 & "' and roid = " & request("roid")
				 set payrs = con.execute(stmt)
				 if not payrs.eof then
				 	do until payrs.eof
				 %>
                 <tr>
                 <td class="style29" style="width: 33%"><%=payrs("ptype")%>
				 &nbsp;</td>
                 <td width="33%" class="style29" style="width: 33%"><%=payrs("pdate")%>
				 &nbsp;</td>
                 <td width="33%" align="right" class="style15" style="width: 33%"><%=formatcurrency(payrs("amt"))%>
				 &nbsp;</td>
				 </tr>
				 <%
				 		payrs.movenext
				 	loop
				 else
				 %>
                 <tr>
                 <td colspan="3" class="style29" style="width: 33%">No Payments posted<br><br><br>
				 &nbsp;</td>
				 </tr>
				 <%
				 end if
				 %>
                 <tr>
                 <td onclick="showPayment()" width="50%" style="cursor:pointer;background-image:url('newimages/rosubheader.jpg'); width: 100%;" class="style26" colspan="3">
				 <br><strong><span class="style13">Add/Delete Payments 
				 (Click)&nbsp;<%
                 paystmt = "select sum(`amt`) as payments from accountpayments where roid = " & request("roid") & " and shopid = '" & request.cookies("shopid") & "'"
                 set payrs = con.execute(paystmt)
                 if isnull(payrs("payments")) then
                 	otherpayments = 0
                 else
                 	otherpayments = payrs("payments")
                 end if
                 response.write formatdollar(otherpayments)
                 set payrs = nothing
                 customerpayments = otherpayments + rs("amtpaid1") + rs("amtpaid2")
                 %></span></strong>&nbsp;<br class="style13"><br></td>
                </tr>
               </table>
</td>
             </tr>
            </table>
           <table width="100%" cellspacing="0" cellpadding="3" class="style16">
             <tr>
              <td style="padding:5px;" width="100%" colspan="2" class="style8">
               <strong>** Totals **</strong></td>
             </tr>
             <%
if taxparts > 0.01 then
	salestax = round(taxparts * (rs("TaxRate") / 100),2)
else
	salestax = 0.00
end if
if ttllbr > 0 then
	labortax = round((ttllbr-rs("discountamt")) * (rs("LaborTaxRate") / 100),2)
else
	labortax = 0
end if

if ttlsub > 0 then
	subtax = round(ttlsub * (rs("SubletTaxRate") / 100),2)
else
	subtax = 0
end if

if rs("totalfees") > 0 then
	feetax = round(rs("totalfees") * (rs("SubletTaxRate") / 100),2)
else
	feetax = 0
end if

nttltax = salestax + labortax + subtax +  feetax
ttlprts = notaxparts + taxparts

             %>
             <tr>
              <td width="50%" align="right" class="style21" style="height: 20px">Total Parts</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15" style="height: 20px"><%=formatdollar(ttlprts)%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" style="height: 22px" class="style21">Total Labor</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" style="height: 22px" class="style15"><%=formatdollar(ttllbr)%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Total Sublet</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(ttlsub)%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td onclick="tempClose()" width="50%" align="right" class="style21">Total Fees</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(rs("TotalFees"))%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Subtotal</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub)%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Tax</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(nttltax)%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Discount</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15">
               <%
               response.write "<input type=hidden name=mysubttl value='" & ttlprts+ttllbr+ttlsub & "'>"
              if rs("DiscountAmt") > 0 then
              	response.write "<font color='red'>(" & formatdollar(rs("DiscountAmt")) & ")</font>"
              else
              	response.write formatdollar(rs("DiscountAmt"))
              end if
              
               %>
               &nbsp;</td>

             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Total RO</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+nttltax-cdbl(rs("DiscountAmt")))%></b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Customer Payments</td>
              <td width="50%" bgcolor="#FFFFFF" style="color:red" align="right" class="style15">&lt; <%=formatdollar(customerpayments)%> &gt;</b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Balance</td>
              <td width="50%" bgcolor="#FFFFFF" style="color:black" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+nttltax-cdbl(rs("DiscountAmt"))-customerpayments)%></b>&nbsp;&nbsp;</td>
             </tr>

           </table>
           </center></center></td>
         </tr>
        </table></td>
      </tr>
      </table>
   </td>
  </tr>
  </table>
  <%
set rs = nothing
strsql = "Select * from repairorders where shopid = '" & request.cookies("shopid") & "' and ROID = " & request("roid")
set nrs = con.execute(strsql)
for fnum = 0 to nrs.fields.count-1
	if nrs.fields(fnum).name = "ROID" then
		fldval = "<input type=hidden name='" & nrs.fields(fnum).name
	else
		fldval = "<input type=hidden name='db" & nrs.fields(fnum).name
	end if
	if nrs.fields(fnum).name = "rodisc" then
		'do nothing
	elseif nrs.fields(fnum).name = "warrdisc" then
		'do nothing
	elseif nrs.fields(fnum).name = "TotalLbr" then
		response.write fldval & "' value='" & ttllbr & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalPrts" then
		response.write fldval & "' value='" & ttlprts & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalSublet" then
		response.write fldval & "' value='" & ttlsub & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "SalesTax" then
		response.write fldval & "' value='" & nttltax & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalRO" then
		response.write fldval & "' value='" & ttllbr+ttlprts+ttlsub+nttltax & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "Subtotal" then
		response.write fldval & "' value='" & ttllbr+ttlprts+ttlsub & "'>" & chr(10)
	else
		response.write fldval & "' value='" & nrs(fnum) & "'>" & chr(10)
	end if
	response.write "<input type=hidden name='fldtype" & nrs.fields(fnum).name
	response.write "' value='" & nrs.fields(fnum).type & "'>" & chr(10)
next

'calc gp
ttls = ttlprts+ttllbr
ttlc = pc+tlc
dprofit = ttls-ttlc
if dprofit > 0 then gp = dprofit/ttls
if gp > 0 then gp = formatpercent(gp)
  %>
<div id="disclosure">
Repair Order Disclosure:<br>
<textarea style="width:100%;height:200px" id="dbrodisc" name="dbrodisc" cols="20" rows="2"><%=rodisc%></textarea><br>
Warranty Disclosure:<br>
<textarea style="width:100%;height:200px" id="dbwarrdisc" name="dbwarrdisc" cols="20" rows="2"><%=warrdisc%></textarea><br>
<input onclick="hideDisclosure()" name="Button1" type="button" value="Close" >
</div>

</form>
</td></tr></table>





<font color="#FFFFFF">


<div id="hider"></div>
<div id="timer">
Saving...<br><br>
<img src="newimages/ajax-loader.gif">
</div>
<div id="estimator"><img style="position:absolute;top:40%;left:45%" src="newimages/ajax-loader.gif"></div>
<form method="post" name="guideform" action="addguidecontents.asp">
<input type="hidden" name="laboritems"><input type="hidden" name="partitems"><input type="hidden" name="comid"><input type="hidden" name="roid" value="<%=request("roid")%>"><input type="hidden" name="shopid" value="<%=request.cookies("shopid")%>">
<div id="estimatorqueue"><b>Labor To be added to RO:</b>  <span style="font-size:x-small">(make any changes you like to the values from the guide, then click on Add All)</span><br></div>
<div id="pestimatorqueue"><b>Parts To be added to RO:</b>  <span style="font-size:x-small">(make any changes you like to the values from the guide, then click on Add All)</span><br></div>
<div id="techselect">
Please select a technician for this labor operation
<select size="1" name="tech">

<%
strsql = "Select EmployeeLast, EmployeeFirst from employees where shopid = '" & request.cookies("shopid") & "' and Active = 'yes'"
set techrs = con.execute(strsql)

if not techrs.eof then
techrs.movefirst
while not techrs.eof
	response.write "<option value='" & techrs ("EmployeeLast") & ", " & techrs ("EmployeeFirst")
	response.write "'>" & techrs ("EmployeeLast") & ", " & techrs ("EmployeeFirst") & "</option>"
	techrs.movenext
wend
end if
set techrs = nothing
%>
 </select><input type="button" value="Select" onclick="">

</div>
</form>
<div id="popuphider"></div>
<div id="popup">
<form name="pmtform" method="post">
<input type="hidden" name="roid" value="<%=roid%>">
<input type="hidden" name="cid" value="<%=cid%>">
<table style="width: 100%" align="center" class="style1" cellspacing="0" cellpadding="3">
	<tr>
		<td style="width: 50%"><strong>Amount:&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input name="amt" type="text" /></td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>How Paid:&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><select name="type">
                   <option value="Cash">Cash</option>
                   <option value="Check">Check</option>
                   <option value="Mastercard">Mastercard</option>
                   <option value="Visa">Visa</option>
                   <option value="American Express">American Express</option>
                   <option value="Discover">Discover</option>
                   <option value="Money Order">Money Order</option>
		</select></td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>Number (optional):&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input name="number" type="text" /></td>
	</tr>
	<tr>
		<td colspan="2" class="style9"><input type="button" onclick="savePayment()" name="b" value="Post Payment">
		<input name="Button1" type="button" onclick="closePayment()" value="Cancel"></td>
	</tr>
</table>
<br>
Other Payments

	<table cellpadding="2" cellspacing="0" style="width: 100%">
		<tr>
			<td style="background-color:maroon;color:white;font-weight:bold">Amount&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">Type&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">Date&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">Delete&nbsp;</td>
		</tr>
		<%
		stmt = "select * from accountpayments where roid = " & request("roid") & " and shopid = '" & request.cookies("shopid") & "'"
		set payrs = con.execute(stmt)
		if not payrs.eof then
			do until payrs.eof
		%>
		<tr>
			<td><%=payrs("amt")%>&nbsp;</td>
			<td><%=payrs("ptype")%>&nbsp;</td>
			<td><%=payrs("pdate")%>&nbsp;</td>
			<td onclick="deletePayment('<%=payrs("id")%>')" style="color:#336699;cursor:pointer">Delete&nbsp;</td>
		</tr>
		<%
				payrs.movenext
			loop
		else
			response.write "<tr><td colspan='4'>No other payments found</td></tr>"
		end if
		%>
	</table>

</form>

</div>
<div id="popupdropshadow"></div>

	<div style="padding:20px;padding-left:40px;width:400px;height:300px;position:absolute;left:45%;top:40%;z-index:999;display:none;background-color:white;border:medium black outset;color:black" id="techfortime<%=crs("laborid")%>">
	<br><br>
	<form name="techfortimeform">
	<input type="hidden" name="roid" value="<%=request("roid")%>">
	<input type="hidden" name="comid">
	<input type="hidden" name="shopid" value="<%=request.cookies("shopid")%>">
	<input type="hidden" name="laborid"><input type="hidden" name="startdate"><input type="hidden" name="enddate">

	Select the technician for this time:<br>
	<select name="techformtimetech">
		<option selected="selected"></option>
		<%
		stmt = "select * from employees where shopid = '" & request.cookies("shopid") & "'"
		'response.write stmt
		set emprs = con.execute(stmt)
		if not emprs.eof then
			do until emprs.eof
		%>
		<option value="<%=emprs("employeelast") & ", " & emprs("employeefirst")%>"><%=emprs("employeelast") & ", " & emprs("employeefirst")%></option>
		<%
				emprs.movenext
			loop
		end if
		%>
	</select><br><br>
	<input onclick="postLaborTimeClock()" name="Button1" type="button" value="Post Time" > <input onclick="cancelTimeClock()" name="Button1" type="button" value="Cancel" >
	</form>
	</div>


<iframe src="rohistorydata.asp?roid=<%=request("roid")%>&v=<%=nrs("vehlicnum")%>" style="display:none" id="history"></iframe>
<iframe frameborder="0" src="upload/upload.asp?roid=<%=request("roid")%>" style="border:medium black outset;width:60%;height:80%;position:absolute;left:20%;top:10%;display:none;z-index:999" id="showframe"></iframe>
<iframe id="picframe" frameborder="0" src='upload/showpics.asp?roid=<%=request("roid")%>' style="border:medium black outset;width:90%;height:90%;position:absolute;left:5%;top:5%;display:none;z-index:999"></iframe>
<div style="border:2px black solid;color:black;font-family:arial;text-align:center;z-index:999;display:none;position:absolute;top:10%;left:20%;width:60%;height:40%;background-color:white" id="alerts">
	<br><strong>You have changed the status of this Repair Order.</strong><br><br>
Would you like to notify the customer via text message or email?

<%
'check for email and cell number and provider

'get email from customer record
stmt = "select * from customer where customerid = " & nrs("customerid") & " and shopid = '" & request.cookies("shopid") & "'"
set ers = con.execute(stmt)
email = ers("email")
set ers = nothing

if len(nrs("cellphone")) > 0 then
	'we have a cell number, do we have a provider?
	
	if len(nrs("cellprovider")) > 0 then
		'we have a cell provider too
		
		celloutput = "<br><br><span onclick='sendUpdate(""text"")' style='color:#336699;font:arial bold;cursor:pointer'>Send Text Message</span>"
		fullcell = "yes"
	else
		celloutput = "<br><br>You have a cell number but no provider.  Please update the customer info to include the cell provider for this feature"
		fullcell = "no"
	end if
else
	celloutput = "<br><br>You have not entered a cell number or provider.  Please update customer information to use this function"
	fullcell = "no"
end if
response.write celloutput

'now check for email address
Function RegExpTest(sEmail)
  RegExpTest = false
  Dim regEx, retVal
  Set regEx = New RegExp

  ' Create regular expression:
  regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"

  ' Set pattern:
  regEx.IgnoreCase = true

  ' Set case sensitivity.
  retVal = regEx.Test(sEmail)

  ' Execute the search test.
  If not retVal Then
    exit function
  End If

  RegExpTest = true
End Function

checkemail = RegExpTest(email)
'response.write checkemail & ":" & email
if checkemail then
	emailoutput = "<br><br><span onclick='sendUpdate(""email"")' style='color:#336699;font:arial bold;cursor:pointer'>Send Email Update</span>"
	fullemail = "yes"
else
	emailoutput = ""
	fullemail = "no"
end if
response.write emailoutput

if fullemail = "yes" and fullcell = "yes" then
	response.write "<br><br><span onclick='sendUpdate(""both"")' style='color:#336699;font:arial bold;cursor:pointer'>Send Both</span>"
end if
response.write "<br><br><span id='closer' onclick='closeUpdateWindow()' style='color:#336699;font:arial bold;cursor:pointer'>Dont Send Update (Close)</span>"
%>

</div>

<script language="javascript">
	status = "Shop Boss Pro"

</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>