<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->

<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<link rel="STYLESHEET" type="text/css" href="css/rich_calendar.css">
<script language="JavaScript" type="text/javascript" src="javascripts/rich_calendar.js"></script>
<script language="JavaScript" type="text/javascript" src="javascripts/rc_lang_en.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/domready.js"></script>
<script type="text/javascript" src="javascripts/killenter.js" ></script>
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




function saveEmp(){
	// check for duplicate pin number
	if (document.theform.pin.value.length > 0){
		xmlHttp=GetXmlHttpObject()
		var url='aspscripts/checkpin.asp?pin='+document.theform.pin.value+'&shopid=<%=request.cookies("shopid")%>'
		trand = parseInt(Math.random()*99999999);  // cache buster
		url = url + "&myrand=" + trand
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
		console.log(url)
		
		xmlHttp.onreadystatechange=function(){
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
				r = xmlHttp.responseText
				console.log("response:"+r)
				if (r == "dup"){
					alert("The PIN number is a duplicate.  Please select a different PIN number")
					return
				}
				if (r == "nodup"){
					sbpaccess = document.getElementById("loginaccess").value
					p = document.getElementById("password").value
					console.log(sbpaccess)
					if(sbpaccess == "Yes" && p == ''){
						alert("All users that are able to login to Shop Boss Pro MUST have a password")
					}else{
						document.theform.submit()
					}
				}
			}
		}
	}else{
		sbpaccess = document.getElementById("loginaccess").value
		p = document.getElementById("password").value
		console.log(sbpaccess)
		if(sbpaccess == "Yes" && p == ''){
			alert("All users that are able to login to Shop Boss Pro MUST have a password")
		}else{
			document.theform.submit()
		}
	}
}
	
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

var xmlHttp

//*** state changed, is the xml response text ready
function stateChanged(){ 
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){ 
		rt = xmlHttp.responseText
		if(rt != "none*none"){
			document.theform.ac.focus()
			tarray = rt.split("*")
			document.getElementById("EmployeeCity").value=tarray[0]
			document.getElementById("EmployeeState").value=tarray[1]
		}else{
			alert("No city or state found.  Possible bad zip code")
		}
	} 
}

function splitIt(splitItStr,splItter){
	myarray = splitItStr.split("*");
	return myarray;
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
	border-color: #000000;
	border-width: 1px;
}
.style3 {
	font-weight: bold;
	border: 0 solid #000000;
}

.style4 {
	border: 0 solid #000000;
	text-align: center;
}

.style5 {
	color: #FFFFFF;
	background-image: url('newimages/pageheader.jpg');
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
}

.style6 {
	font-weight: bold;
	border-width: 1px;
}

.style7 {
	border: 0 solid #000000;
}

a.tooltip {outline:none; }
a.tooltip strong {line-height:30px;}
a.tooltip:hover {text-decoration:none;} 
a.tooltip span {
    z-index:1010;display:none; padding:14px 20px;
    margin-top:-30px; margin-left:-285px;
    width:240px; line-height:16px;
}
a.tooltip:hover span{
    display:inline; position:absolute; color:#111;
    border:1px solid #DCA; background:#fffAF0;}
.callout {z-index:1020;position:absolute;top:30px;border:0;left:-12px;}
    
/*CSS3 extras*/
a.tooltip span
{
    border-radius:4px;
    -moz-border-radius: 4px;
    -webkit-border-radius: 4px;
        
    -moz-box-shadow: 5px 5px 8px #CCC;
    -webkit-box-shadow: 5px 5px 8px #CCC;
    box-shadow: 5px 5px 8px #CCC;
}

.style8 {
	font-size: xx-small;
}

-->
</style>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<input type="hidden" name="dfield" id="dfield">
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
                <form name=theform action=addempnow.asp method=post>
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                      <td height="11">
                       <p align="center">
                       <%
                       if len(request("nodate")) > 1 then response.write "You must enter the date hired"
                       %>
                       <input type="button" onclick="location.href='employees.asp'" value="Employee List" name="Abutton1">
						<a href="employees.asp"><br>
                       </a>
                       <b>Complete information about your employee.&nbsp; All Fields are required.</b>
                       </p>
                        <table border=2 cellpadding=0 cellspacing=0 width="80%" align=center>
                       <tr>
                       <td style="height:30; background-image:url('newimages/pageheader.jpg')" class="style6">
                        <p align="center"><font face="Arial" color="#FFFFFF" size="3">Employee
                        Information</font></p>
                       </td>
                       <tr>
                       <td class="style1" style="width: 25%">
					   <table border="0" cellspacing="0" cellpadding="2" style="width: 100%">
                        <tr>
                         <td align="right" style="width: 25%" class="style3">Last
                          Name</td>
                         <td style="width: 25%" class="style7">
						 <input id="EmployeeLast" title="Last Name" type="text" name="EmployeeLast" size="25" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; tabindex="2" value="<%=request("EmployeeLast")%>" tabindex="1"></td>
                         <td style="width: 25%" class="style3">
                          Position</td>
                         <td style="width: 25%" class="style3">
						 <select id="Position" title="Position" size="1" name="Position" style="font-variant: small-caps; font-family: Arial; font-size: 10pt; tabindex="11" tabindex="4">
                         
                         <%
                         jstmt = "select * from jobdesc where shopid = '" & request.cookies("shopid") & "'"
                         set jrs = con.execute(jstmt)
                         if not jrs.eof then
                         	jrs.movefirst
                         	while not jrs.eof
                         		response.write "<option value='" & jrs("JobDesc") & "'>" & jrs("JobDesc") & "</option>" & chr(10)
                         		jrs.movenext
                         	wend
                         end if
                         %>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" style="width: 25%; height: 24px;" class="style3">First
                          Name</td>
                         <td style="width: 25%; height: 24px;" class="style7">
						 <input id="EmployeeFirst" title="First Name" name="EmployeeFirst" size="25" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; tabindex="3" value="<%=request("EmployeeFirst")%>" tabindex="2"></td>
                         <td style="width: 25%; height: 24px;" class="style3">
                          Default
                          Service Writer</td>
                         <td style="width: 25%; height: 24px;" class="style7">
						 <select size="1" name="DefaultWriter" style="font-family: Arial; font-variant: small-caps; font-size: 10pt" tabindex="5">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" style="width: 25%" class="style3">Password</td>
                         <td style="width: 25%" class="style7">
                         <input id="password" title="Password" type="text" name="password" size="25" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; value="<%=mypassword%>" tabindex="3"></td>
                         <td style="width: 25%" class="style3">
                          Hourly Rate</td>
                         <td style="width: 25%" class="style3"><b>$</b><input id="hourlyrate" title="Flat Rate/Hourly Rate" type="text" name="hourlyrate" size="6" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; height: 20px;"17" value="0" tabindex="6"></td>
                        </tr>
                        <tr>
                         <td align="right" style="width: 25%" class="style3">Timeclock PIN Number</td>
                         <td style="width: 25%" class="style7">
                         <input id="pin" title="Password" type="text" name="pin" size="25" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; value="<%=mypassword%>" tabindex="3"></td>
                         <td style="width: 25%" class="style3">
                          Pay Type</td>
                         <td style="width: 25%" class="style3">						 
                         <select name="paytype">
						 	<option value="flatrate">Flat Rate</option>
						 	<option value="hourly">Hourly</option>
						 </select>
</td>
                        </tr>
                        <tr>
                         <td align="right" style="width: 25%" class="style3" valign="top">Michigan Mechanic #</td>
                         <td style="width: 25%" class="style7" valign="top">
                         <input id="mechanicnumber" title="Password" type="text" name="mechanicnumber" size="25" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; value="<%=mypassword%>" tabindex="3"></td>
                         <a href="#" class="tooltip">
						<span>
						Find/View and Edit your customers and their vehicles.
						</span>
						</a>
                         <td style="width: 25%" class="style3" valign="top">
                          Mode - Full or Tech<br><span class="style8">(Tech mode 
						  shows a technicial only their jobs, instead of all 
						  jobs.&nbsp; It also only has the functions they need 
						  to work on a vehicle)</span></td>
                         <td style="width: 25%" class="style3" valign="top">
						 <select name="mode">
						 <option value="Full">Full</option>
						 <option value="Tech">Tech</option>
						 </select></td>
                        </tr>
                        <tr>
                         <td align="right" style="width: 25%" class="style3" valign="top">
						 Employee is a Technician/Mechanic</td>
                         <td style="width: 25%" class="style7" valign="top">
                         <select name="showtechlist">
						 <option value="yes">Yes</option>
						 <option value="no">No</option>
						 </select></td>
                         <td style="width: 25%" class="style3" valign="top">
                          Date Hired</td>
                         <td style="width: 25%" class="style3" valign="top"><input type="text" id="sd" name="sd" value="" 
						style="width: 100px" >
						 &nbsp;</td>
                        </tr>
                        </table>
                        <table style="width:100%" cellpadding="3" cellspacing="0">
                        <tr>
                         <td align="center" colspan="5" height="16" style="width: 100%" class="style5">
                         <strong>Login Access</strong></td>
                        </tr>
                        <%
                        %>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Login to Shop Boss Pro</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="loginaccess" id="loginaccess" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="18">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Edit Inventory</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="EditInventory" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="27">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Settings Access</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="CompanyAccess" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="19">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Re-Open RO</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ReOpenRO" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="28">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Employee Access</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="EmployeeAccess" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="20">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Change User Security</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangeUserSecurity" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="29">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Report Access</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ReportAccess" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="21">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Change Parts Matrix</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangePartMatrix" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="30">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Create RO</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="CreateRO" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="22">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Change Parts Codes</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangePartCodes" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="31">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Create Part Sale Invoice</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="CreateCT" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="23">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Change Job Descriptions</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangeJobDescription" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="32">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Edit Supplier</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="EditSupplier" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="24">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Change Sources</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangeSources" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="33">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Inventory Lookup</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="InventoryLookup" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="25">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Accounting Access</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="accounting" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="33">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Change RO Types</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="ChangeRepairOrderTypes" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="34">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         Delete Customer?</td>
                         <td style="width: 25%; height: 16%;" class="style7">
							<select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="deletecustomer" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="33">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select>
                         </td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Delete Parts and Labor from RO</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="candelete" style="font-family: Arial;font-variant:small-caps; font-size: 10pt" tabindex="34">
                           <option value="yes">Yes</option>
                           <option value="no">No</option>
                          </select></td>
                         <td style="width: 25%; height: 16%;" class="style7" colspan="2">
                         &nbsp;</td>
                         <td style="width: 25%; height: 16%;" class="style7">
                         &nbsp;</td>
                        </tr>
                        <tr>
                         <td align="right" class="style7" style="width: 25%; height: 16%">
                         Change Shop Alerts &amp; Notices</td>
                         <td colspan="2" class="style7" style="width: 25%; height: 16%">
                         <select size="1" name="changeshopnotice" style="font-family: Arial;font-variant:small-caps; font-size: 10pt" tabindex="34">
                           <option value="yes">Yes</option>
                           <option value="no">No</option>
                          </select>&nbsp;</td>
                         <td class="style7" style="width: 25%; height: 16%">
                         ToolTips? (Tooltips are pop ups that explain certain 
						 buttons and functions for new users)</td>
                         <td class="style7" style="width: 25%; height: 16%">
                         <select onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" size="1" name="tooltip" style="font-variant: small-caps; font-family: Arial; font-size: 10pt" tabindex="33">
                           <option value="Yes">Yes</option>
                           <option value="No">No</option>
                          </select></td>
                        </tr>
                        <tr>
                         <td colspan="5" style="width: 100%" class="style4">
                          <input onclick="saveEmp()" name="Submit1" type="button" value="Save">
						  <input onclick="location.href='employees.asp'" name="Button1" type="button" value="Cancel"></td>
                        </tr>
                       </table>
                       </table></td>
                    </tr>
                  </table>
                  </form>
                </div>
                &nbsp;</td>
            </tr>
            <tr>
              <td width="110" valign="top" align="left" height="1"></td>
              <td valign="top" width="100%"></td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
</div>
<p>&nbsp;</p>

</body>
<link href="jquery/jquery-ui.css" rel="stylesheet" >
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.0/jquery.min.js"></script>
<script src="javascripts/jquery-ui.js"></script>

<script language="javascript">
$('#sd').datepicker();
$("#sd").attr('readonly', 'readonly');

</script>

<script language="javascript">
document.theform.EmployeeLast.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>