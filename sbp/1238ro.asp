<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
' just adding a comment
shopid = request.cookies("shopid")
shopname = request.cookies("shopname")

stmt = "select * from companyadds where `name` = 'WorldPac Integrated Parts Ordering' and shopid = '" & shopid & "'"
set rs = con.execute(stmt)
if rs.eof then
	worldpac = "no"
else
	worldpac = "yes"
end if

stmt = "Select * from company where shopid = '" & shopid & "'"
set nrs = con.execute(stmt)
showgp = lcase(nrs("showgp"))
npuser = nrs("nexpartusername")
nppwd = nrs("nexpartpassword")

RequireSource = lcase(nrs("RequireSource"))
RequireComplaint = lcase(nrs("RequireComplaint"))
requireoutmileage = lcase(nrs("requireoutmileage"))
requiretirepressure = lcase(nrs("requiretirepressure"))
merchantaccount = nrs("merchantaccount")
merchantaccounttype = nrs("merchantaccounttype")
userfee1max = nrs("userfee1max")
userfee2max = nrs("userfee2max")
userfee3max = nrs("userfee3max")
milesinlabel = nrs("milesinlabel")
milesoutlabel = nrs("milesoutlabel")
showropayments = nrs("showropayments")
chargeshopfeeson = nrs("chargeshopfeeson")
userfee1taxable = nrs("userfee1taxable")
userfee2taxable = nrs("userfee2taxable")
userfee3taxable = nrs("userfee3taxable")
requiremileagetoprintro = nrs("requiremileagetoprintro")
requiretirepressuretoprintro = nrs("requiretirepressuretoprintro")
printpopup = nrs("printpopup")
customropage = nrs("customropage")
carfax = nrs("carfaxlocation")
esign = lcase(nrs("esign"))
useim = lcase(nrs("useim"))
if lcase(customropage) <> "no" then 
	tredir = customropage & "?roid=" & request("roid") & "&sp=" & request("sp")
	'response.redirect tredir
end if
if esign = "no" then
	'response.redirect "ro0.asp?" & request.servervariables("query_string")
end if

%>

<html>

<head><meta name="robots" content="noindex,nofollow"/>
<style type="text/css">
<!--
a            { text-decoration: none }
-->
select{
	width:150px;
}
   code, pre {font-family: Consolas,monospace; color: green;}
   ol li {margin: 0 0 15px 0;}
#epicorselectbutton:hover{
	background-color:#E8E8E8;
}
#worldpacselectbutton:hover{
	background-color:#E8E8E8;
}
#nexpartselectbutton:hover{
	background-color:#E8E8E8;
}

</style>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0"/>
<meta name="ProgId" content="FrontPage.Editor.Document"/>
<script type="text/javascript" src="jquery/jquery-1.10.2.min.js"></script>
<title><%=shopname%></title>
<%
'on error resume next
function formatdollar(a)

	if isnull(a) then
		formatdollar = "0.00"
	else	
		formatdollar = formatnumber(a,2)
	end if

end function
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
'check for required fields
strsql = "Select * from repairorders where shopid = '" & shopid & "' and ROID = " & request("roid")



'response.write strsql
set rs = con.execute(strsql)
origshopid = request("origshopid")
currshopid = rs("shopid")

rostat = rs("status")
vehid = rs("vehid")
if ucase(rs("Status")) = "CLOSED" then 
	sstat = "closed"
	if len(origshopid) > 0 then 
		qs = "origshopid=" & origshopid & "&"
	else
		qs = ""
	end if
	if request("sp") = "history.asp" then
		newsp = "viewro.asp?" & qs & "lic=" & request("lic") & "&roid=" & request("roid") & "&sp=" & request("sp")
	elseif request("sp") = "findroframe.asp" then
		newsp = "viewro.asp?" & qs & "roid=" & request("roid") '& "&srch=" & request("srch") & "&sf=" & request("sf") & "&SortBy=" & request("SortBy") & "&sp=" & request("sp")
	elseif request("sp") = "wip.asp" then
		newsp = "viewro.asp?" & qs & "roid=" & request("roid")
	elseif request("sp") = "viewro.asp" then
		newsp = "viewro.asp?" & qs & "roid=" & request("roid")
	else
		newsp = "wip.asp"
	end if
	'response.write newsp
	response.redirect newsp
end if


rodisc = rs("rodisc")
if RegExpTest(rs("email")) then
	customeremail = rs("email")
else
	customeremail = ""
end if
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

response.write chr(10) & "<script language=""javascript"">" & chr(10)
response.write vbTab & "function saveAll(dir, itype, sid, pid, lid){" & chr(10)
response.write vbTab & "//check for miles, writer and source" & chr(10)
response.write vbTab & "i = 0" & chr(10)
response.write vbTab & "newline = '\n \n'" & chr(10)
response.write vbTab & "errmess = 'Required Fields are not Complete'" & chr(10)


if ucase(rs("Status")) = "CLOSED" or ucase(rs("Status")) = "FINAL" then
	if RequireSource = "yes" then
		response.write vbTab & "if (document.theform.Source.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Source is a required field.'}" & chr(10)
	end if
	if RequireComplaint = "yes" then
		response.write vbTab & "if (document.theform.Comments.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Comments is a required field.'}" & chr(10)
	end if
	if lcase(requireoutmileage) = "yes" then
		response.write vbTab & "if (document.theform.MilesOut.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Vehicle In Miles is required'}" & chr(10)
	end if
	if lcase(requireoutmileage) = "yes" then
		response.write vbTab & "if (document.theform.VehicleMiles.value.length == 0){" & chr(10)
		response.write vbTab & vbTab & "i++" & chr(10)
		response.write vbTab & vbTab & "errmess = errmess + newline + 'Vehicle Out Miles is required'}" & chr(10)
	end if
end if	

if lcase(requiretirepressure) = "yes" then
	response.write "if (document.theform.Status.value == 'CLOSED' || document.theform.Status.value == 'FINAL'){" & chr(10)
	response.write vbTab & "tpinlf = document.theform.tirepressureinlf.value.length" & chr(10)
	response.write vbTab & "tpinlr = document.theform.tirepressureinlr.value.length" & chr(10)
	response.write vbTab & "tpinrf = document.theform.tirepressureinrf.value.length" & chr(10)
	response.write vbTab & "tpinrr = document.theform.tirepressureinrr.value.length" & chr(10)
	
	response.write vbTab & "tpoutlf = document.theform.tirepressureoutlf.value.length" & chr(10)
	response.write vbTab & "tpoutlr = document.theform.tirepressureoutlr.value.length" & chr(10)
	response.write vbTab & "tpoutrf = document.theform.tirepressureoutrf.value.length" & chr(10)
	response.write vbTab & "tpoutrr = document.theform.tirepressureoutrr.value.length" & chr(10)

	response.write vbTab & "tdlf = document.theform.treaddepthlf.value.length" & chr(10)
	response.write vbTab & "tdlr = document.theform.treaddepthlf.value.length" & chr(10)
	response.write vbTab & "tdrf = document.theform.treaddepthlf.value.length" & chr(10)
	response.write vbTab & "tdrr = document.theform.treaddepthlf.value.length" & chr(10)
	
	response.write vbTab & "if (tpinlf==0||tpinlr==0||tpinrf==0||tpinrr==0||tpoutlf==0||tpoutlr==0||tpoutrf==0||tpoutrr==0||tdlf==0||tdlr==0||tdrf==0||tdrr==0){" & chr(10)
	response.write vbTab & "i++" & chr(10)
	response.write vbTab & "errmess = errmess + newline + 'Tire Pressures and Tread Depth are required'}" & chr(10)
	response.write "}" & chr(10)
end if
		

response.write vbTab & "if (i > 0){" & chr(10)
response.write vbTab & vbTab & "alert(errmess+newline)" & chr(10)
response.write vbTab & "}else{" & chr(10)
response.write vbtab & "if(document.getElementById('timerunning').value == 'no'){" & chr(10)
response.write vbTab & "document.getElementById('hider').style.display='block'" & chr(10)
response.write vbTab & "document.getElementById('timer').style.display='block'" & chr(10)
response.write vbTab & "document.theform.dir.value = dir" & chr(10)
response.write vbTab & "document.theform.itype.value = itype" & chr(10)
response.write vbTab & "document.theform.sid.value = sid" & chr(10)
response.write vbTab & "document.theform.pid.value = pid" & chr(10)
response.write vbTab & "document.theform.lid.value = lid" & chr(10)
response.write vbtab & "calcPercents()" & chr(10)
response.write vbTab & "document.theform.submit();" & chr(10)
response.write vbtab & "}" & chr(10)
response.write vbtab & "setTimeout('closeTimer()',5000)}}" & chr(10)
response.write "</script>"

homephone = "(" & left(rs("CustomerPhone"),3) & ") " & mid(rs("CustomerPhone"),4,3) & "-" & right(rs("CustomerPhone"),4)
workphone = "(" & left(rs("CustomerWork"),3) & ") " & mid(rs("CustomerWork"),4,3) & "-" & right(rs("CustomerWork"),4)
cellphone = "(" & left(rs("CellPhone"),3) & ") " & mid(rs("CellPhone"),4,3) & "-" & right(rs("CellPhone"),4)

'**************************************************************************
'****** ttllbr, ttlprts, ttlsub all come from complaintlayout.asp *********
'**************************************************************************

cemail = rs("email")
%>

<script type="text/javascript" src="javascripts/ro.js"></script>
<link rel="STYLESHEET" type="text/css" href="css/rich_calendar.css">
<script language="JavaScript" type="text/javascript" src="javascripts/rich_calendar.js"></script>
<script language="JavaScript" type="text/javascript" src="javascripts/rc_lang_en.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/domready.js"></script>
<script type="text/javascript">
function discountPercent(e){return isNumber(e)?isNumber(e)&&isNumber(document.theform.discsubtotal.value)?(document.theform.DiscountAmt.value=Math.round(e/100*document.theform.discsubtotal.value),chgDB("DiscountAmt"),saveAll("ro.asp?a=n&","","","","",""),void 0):void alert("You must have a subtotal amount for the RO before you can give a discount"):void alert("Discount percentage must be zero or greater")}function show_cal(e,t){if(!cal_obj2){document.getElementById("dfield").value=t;var l=document.getElementById(document.getElementById("dfield").value);l.disabled=!0,cal_obj2=new RichCalendar,cal_obj2.start_week_day=0,cal_obj2.show_time=!1,cal_obj2.language="en",cal_obj2.user_onchange_handler=cal2_on_change,cal_obj2.user_onclose_handler=cal2_on_close,cal_obj2.user_onautoclose_handler=cal2_on_autoclose,cal_obj2.parse_date(l.value,format),cal_obj2.show_at_element(l,"adj_right-")}}function cal2_on_change(e,t){"day"==t&&(document.getElementById(document.getElementById("dfield").value).disabled=!1,document.getElementById(document.getElementById("dfield").value).value=e.get_formatted_date(format),("StatusDate"==document.getElementById("caldatefield").value||"DateIn"==document.getElementById("caldatefield").value)&&chgDB(document.getElementById("caldatefield").value),e.hide(),cal_obj2=null)}function cal2_on_close(e){document.getElementById(document.getElementById("dfield").value).disabled=!1,e.hide(),cal_obj2=null}function cal2_on_autoclose(){document.getElementById(document.getElementById("dfield").value).disabled=!1,cal_obj2=null}var cal_obj2=null,format="%m/%d/%Y";
</script>
<style type="text/css">
li,p,td,th{font-size:10pt;font-family:Verdana,Arial,Helvetica}.style1{border-width:0;font-family:Arial;color:#fff;background-color:#06C}.abuts{font-size:12px;width:100px;height:24px;cursor:hand;color:#000}#hider{position:absolute;top:0;left:0;width:100%;height:120%;background-color:gray;-ms-filter:"alpha(Opacity=70)";filter:alpha(opacity=70);-moz-opacity:.7;opacity:.7;z-index:998;display:none}#timer{background-color:#fff;position:absolute;top:25%;left:35%;width:30%;height:20%;z-index:999;display:none;text-align:center;border:4px navy outset;color:#000;padding-top:2%;font-family:Verdana,Arial,sans-serif;font-weight:700}#cctimer{background-color:#fff;position:absolute;top:25%;left:35%;width:30%;height:20%;z-index:1050;display:none;text-align:center;border:4px navy outset;color:#000;padding-top:2%;font-family:Verdana,Arial,sans-serif;font-weight:700}#estimator{background-color:#fff;position:absolute;top:1%;left:1%;width:97%;height:95%;z-index:999;display:none;border:4px navy outset;color:#000;padding:5px;overflow:auto}#estimatorqueue{background-color:#fff;position:absolute;top:1%;left:1%;width:48%;height:16%;z-index:999;display:none;border:4px navy outset;color:#000;padding:5px;overflow:auto}#techselect{background-color:#fff;position:absolute;top:20%;left:30%;width:40%;height:25%;z-index:1000;display:none;border:4px navy outset;color:#000;padding:5px;padding-top:30px;overflow:auto}#pestimatorqueue{background-color:#fff;position:absolute;top:1%;right:1%;width:48%;height:16%;z-index:999;display:none;border:4px navy outset;color:#000;padding:5px;overflow:auto}#history{width:90%;height:90%;top:5%;left:5%;background-color:#fff;position:absolute;z-index:1000;border:medium navy outset}#inspection{width:90%;height:90%;top:60px;left:5%;background-color:#fff;position:absolute;z-index:1000;border:medium navy outset}#inspectionbutton{top:0;left:5%;width:300px;height:50px;font-size:x-large;position:absolute;z-index:1002;color:red;font-weight:700}--> .menustyle{text-align:center;background-color:#F0F0F0;width:7.7%;border:1px gray outset;color:maroon;font-weight:700;cursor:pointer;font-size:x-small}.style8{text-align:center;color:#FFF;border-width:1px;background-image:url(newimages/pageheader.jpg)}.style9{color:#FFF;font-weight:700;border-width:1px;background-image:url(newimages/rosubheader.jpg)}.style10{background-image:url(newimages/wipheader.jpg);color:#fff;font-weight:700}.style13{color:#FFF}.style14{color:#fff;border-width:0;background-color:#06C}.style15{border-width:0}.style16{border-collapse:collapse;border-style:solid;border-width:1px}.style17{border-style:solid;border-width:0;color:#FFF;background-color:#06C}.style18{text-align:right}.style20{border-width:0;text-align:right}.style21{color:#FFF;border-width:1px;background-color:#06C}body,div,p,span{font-family:Arial,Helvetica,sans-serif}#popup{position:absolute;top:100px;left:35%;width:400px;min-height:200px;overflow-y:auto;border:medium navy outset;text-align:center;color:#000;display:none;z-index:999;background-color:#fff;padding:20px}#popuphider{position:absolute;top:0;left:0;width:100%;height:120%;background-color:gray;-ms-filter:"alpha(Opacity=50)";filter:alpha(opacity=70);-moz-opacity:.7;opacity:.7;z-index:997;display:none}#cchider{position:absolute;top:0;left:0;width:100%;height:120%;background-color:gray;-ms-filter:"alpha(Opacity=50)";filter:alpha(opacity=70);-moz-opacity:.7;opacity:.7;z-index:1049;display:none}.style22{font-size:x-small}#recrepairs{position:absolute;top:100px;left:10%;width:80%;height:500px;overflow-y:auto;border:medium navy outset;text-align:center;color:#000;display:block;z-index:1150;background-color:#fff;padding:20px;font-size:medium;font-weight:700}#addrecrepairs{position:absolute;top:100px;left:15%;width:60%;height:500px;overflow-y:auto;border:medium navy outset;text-align:center;color:#000;display:none;z-index:1050;background-color:#fff;padding:20px;font-size:medium;font-weight:700}#editrecrepairs{position:absolute;top:100px;left:15%;width:60%;height:500px;overflow-y:auto;border:medium navy outset;text-align:center;color:#000;display:none;z-index:1150;background-color:#fff;padding:20px;font-size:medium;font-weight:700}#disclosure{position:absolute;top:100px;left:15%;width:60%;height:500px;overflow-y:auto;border:medium navy outset;text-align:center;color:#000;display:none;z-index:999;background-color:#fff;padding:20px;font-size:medium;font-weight:700}#customvehiclefields{position:absolute;top:100px;left:20%;width:60%;height:500px;overflow-y:auto;border:medium navy outset;color:#000;display:none;z-index:1005;background-color:#fff;padding:20px;font-size:medium;font-weight:700;text-align:left}#spouseinfo{position:absolute;top:100px;left:30%;width:40%;height:500px;overflow-y:auto;border:medium navy outset;color:#000;display:none;z-index:1005;background-color:#fff;padding:20px;font-size:medium;font-weight:700;text-align:left}.style24{border-style:solid;border-width:0;text-align:center;color:#FFF;background-color:#060}.style25{border-style:solid;border-width:0;color:#FFF;background-color:#06C;text-align:center}.style26{color:#fff;border-width:0;background-color:#06C;text-align:center}.style27{color:#fff;border-width:0;background-color:#06C;text-align:left}.style29{border-width:0;text-align:left}.style30{border-style:solid;border-width:0;color:#FFF;background-color:#06C;text-align:right}.style31{font-size:x-small;color:#FFF;background-color:#06C}.style32{color:#FFF;background-color:#06C}.style33{background-color:#060}.style34{background-image:url(newimages/wipheader.jpg);color:#fff;font-weight:700;text-align:center}.style35{font-size:xx-small}.style36{text-align:center}#component_selection,.admin,.branding,.car_selection,.component_search,.user{display:none}h1{font-size:14px}.style37{background-color:#FFC}.style38{background-color:#CFC}.style39{background-color:#CFF}
</style>

<script type="text/javascript" language="javascript">
function closeTimer(){document.getElementById("timer").style.display="none",document.getElementById("hider").style.display="none"}function showHistory(){document.getElementById("hider").style.display="block",document.getElementById("history").style.display="block"}function reloadPage(){document.body.onunload="",saveAll("ro.asp?a=n&","","","","","")}function chgDB(fldchg){var dbchg="db"+fldchg;docitem="document.theform."+fldchg,docval=eval(docitem+".value"),newdocval=replace(replace(docval,"\r"," "),"\n"," "),fldval="document.theform."+dbchg+".value='"+newdocval+"'",eval(fldval),document.getElementById("changes").value="yes"}function closeRO(){if("CLOSED"==document.theform.Status.value)if(c=confirm("This will close the RO.  Click OK to proceed or Cancel to Cancel")){document.body.onunload='';document.theform.Status.focus();var e=(new Date,new Date),t=e.getMonth()+1,n=e.getDate(),o=e.getFullYear(),r=t+"/"+n+"/"+o;"no"==document.theform.dboverridestatusdate.value&&(document.theform.dbFinalDate.value=r,document.theform.dbStatusDate.value=r),saveAll("wip.asp?a=n&","","","","","")}else document.theform.Status.value="Final"}function replace(e,t,n){var o=e.length,r=t.length;if(0==o||0==r)return e;var a=e.indexOf(t);if(!a&&t!=e.substring(0,r))return e;if(-1==a)return e;var l=e.substring(0,a)+n;return o>a+r&&(l+=replace(e.substring(a+r,o),t,n)),l}function isInteger(e){var t;for(t=0;t<e.length;t++){var n=e.charAt(t);if("0">n||n>"9")return!1}return!0}function stripCharsInBag(e,t){var n,o="";for(n=0;n<e.length;n++){var r=e.charAt(n);-1==t.indexOf(r)&&(o+=r)}return o}function daysInFebruary(e){return e%4!=0||e%100==0&&e%400!=0?28:29}function DaysArray(e){for(var t=1;e>=t;t++)this[t]=31,(4==t||6==t||9==t||11==t)&&(this[t]=30),2==t&&(this[t]=29);return this}function isDate(e){var t=DaysArray(12),n=e.indexOf(dtCh),o=e.indexOf(dtCh,n+1),r=e.substring(0,n),a=e.substring(n+1,o),l=e.substring(o+1);strYr=l,"0"==a.charAt(0)&&a.length>1&&(a=a.substring(1)),"0"==r.charAt(0)&&r.length>1&&(r=r.substring(1));for(var d=1;3>=d;d++)"0"==strYr.charAt(0)&&strYr.length>1&&(strYr=strYr.substring(1));return month=parseInt(r),day=parseInt(a),year=parseInt(strYr),-1==n||-1==o?!1:r.length<1||1>month||month>12?!1:a.length<1||1>day||day>31||2==month&&day>daysInFebruary(year)||day>t[month]?!1:4!=l.length||0==year||minYear>year||year>maxYear?!1:-1!=e.indexOf(dtCh,o+1)||0==isInteger(stripCharsInBag(e,dtCh))?!1:!0}function ValidateForm(){var e=document.frmSample.txtDate;return 0==isDate(e.value)?(e.focus(),!1):!0}function setStatusDate(){var e=new Date,t=e.getDate(),n=10>t?"0"+t:t,o=e.getMonth()+1,r=10>o?"0"+o:o,a=e.getYear(),l=1e3>a?a+1900:a,d=r+"/"+n+"/"+l;"no"==document.theform.dboverridestatusdate.value&&(document.theform.dbFinalDate.value=d,document.theform.dbStatusDate.value=d)}function changeColor(e){(document.getElementById(e).style.backgroundColor="white")?(document.getElementById(e).style.backgroundColor="green",document.getElementById(e).style.backgroundColor="white"):(document.getElementById(e).style.backgroundColor="white",document.getElementById(e).style.backgroundColor="#000080")}function uploadFiles(){document.getElementById("showframe").style.display="block",document.getElementById("hider").style.display="block"}function showPics(){document.getElementById("picframe").style.display="block",document.getElementById("hider").style.display="block"}function closeUpdateWindow(){return"y"==document.theform.x.value?(document.getElementById("hider").style.display="none",document.getElementById("alerts").style.display="none",void(document.theform.x.value="")):void(t=setTimeout("document.theform.submit()",500))}function showUpdateWindow(e){return"y"==e?(document.getElementById("hider").style.display="block",document.getElementById("alerts").style.display="block",void(document.theform.x.value="y")):void(document.theform.Status.value!=document.theform.oldstatus.value?(document.getElementById("hider").style.display="block",document.getElementById("alerts").style.display="block"):t=setTimeout("document.theform.submit()",500))}function stateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(alert(xmlHttp.responseText),closeUpdateWindow())}function GetXmlHttpObject(){var e=null;return window.XMLHttpRequest?e=new XMLHttpRequest:window.ActiveXObject&&(e=new ActiveXObject("Microsoft.XMLHTTP")),e}var xmlHttp,dtCh="/",minYear=1900,maxYear=2100,xmlHttp;
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

function closeGuide(){document.getElementById("estimator").style.display="none",document.getElementById("hider").style.display="none",document.getElementById("estimatorqueue").style.display="none",document.getElementById("pestimatorqueue").style.display="none"}
var r, c

function showGuide(roid,comid,v){

	document.getElementById("estimator").innerHTML = '<div style="position:absolute;top:40%;left:0%;font-size:18px;text-align:center;padding-left:40%">Loading...<br>(images may increase loading time)<br><img style="" src="newimages/ajax-loader.gif"></div>'
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
			alldataurl = "/v1/vins/"+vin
		}else{
			alldataurl = "/v1/years"
		}
	}
	if (v == "n"){
		alldataurl = "/v1/years"
	}
	//alert(alldataurl)
	var url="guide/alldatadisplay.asp?vin="+vin+"&roid=<%=request("roid")%>&comid="+comid+"&r="+r+"&alldatashopid=<%=shopid%>&alldataurl="+alldataurl
	//alert(url)
	guideTimer = setTimeout("cancelGuide()",15000)
	xmlHttp.onreadystatechange=guideStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function guideStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,clearTimeout(guideTimer),"bad vin"==rt?(sg="showGuide('"+r+"','"+c+"','n')",eval(sg)):setTimeout(function(){loadGuide(rt)},100))}
function cancelGuide(){alert("Labor Guide may be unavailable.  Please try again.  You may need to close your browser and re-open Shop Boss Pro"),document.getElementById("estimator").innerHTML="",document.getElementById("estimator").style.display="none",document.getElementById("hider").style.display="none",clearTimeout(guideTimer)}function loadGuide(e){"none"!=e?(document.getElementById("estimator").innerHTML=e,document.getElementById("lastguide").value=e):document.getElementById("estimator").innerHTML=document.getElementById("lastguide").value}
function getLink(shopid,url,roid,comid){

	document.getElementById("estimator").innerHTML = '<div style="position:absolute;top:40%;left:0%;font-size:18px;text-align:center;padding-left:40%">Loading...<br>(images may increase loading time)<br><img style="" src="newimages/ajax-loader.gif"></div>'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	if (document.getElementById("goback").value == "yes"){
		if (document.getElementById("gobackurl").value.length > 0){
			url = document.getElementById("gobackurl").value
		}
	}
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	if (url.length > 0){
		document.getElementById("gobackurl").value = url
	}
	var url="guide/alldatadisplay.asp?roid=<%=request("roid")%>&comid="+document.guideform.comid.value+"&r="+r+"&alldatashopid=<%=shopid%>&alldataurl="+url
	//alert(url)
	xmlHttp.onreadystatechange=guideStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function getSearch(shopid,urla,urlb,roid,comid,searchStr){
	
	document.getElementById("estimator").innerHTML = '<div style="position:absolute;top:40%;left:0%;font-size:18px;text-align:center;padding-left:40%">Loading...<br>(images may increase loading time)<br><img style="" src="newimages/ajax-loader.gif"></div>'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	//alert(urla)
	vin = document.theform.dbVin.value
	r=Math.floor(Math.random()*111111)
	//searchStr = document.getElementById("alldatasearchfor").value
	alldataurl = urla + encodeURIComponent(searchStr.replace(/&amp;/g, "&")) + urlb
	//alert(alldataurl)
	var url="guide/alldatadisplay.asp?roid=<%=request("roid")%>&comid="+document.guideform.comid.value+"&r="+r+"&alldatashopid=<%=shopid%>&alldataurl="+alldataurl
	//alert(url)
	xmlHttp.onreadystatechange=guideStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}


function selectTech(){document.getElementById("techselect").style.display="block"}function saveTech(e,c){"select"!=t?document.getElementById("labor*"+c+"*de").value=e:alert("Please select a technician for this labor operation")}
function addLabor(operationName,componentName,componentItem,laborTime){

	//create a tech dropdown for the labor added queue
	<%
	techbuild = "<select style='font-size:10px;' size='1' id='tech' name='tech'>"
	strsql = "Select EmployeeLast, EmployeeFirst from employees where shopid = '" & shopid & "' and Active = 'yes'"
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
	randomnumber = makeid()
	descriptionname = removeSpaces(operationName+componentName+componentItem+randomnumber)
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

function makeid(){for(var e="",t="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",n=0;10>n;n++)e+=t.charAt(Math.floor(Math.random()*t.length));return e}function addPart(e,t,n,l){newline="<span style='font-size:12px;font-weight:bold;'>#: </span><input style='font-size:10px;' name='part*"+t+"*pn' size='20' type='text' name='"+t+"pn' value='"+t+"'>",newline+="<span style='font-size:12px;font-weight:bold;'> Desc: </span><input style='font-size:10px;' name='part*"+t+"*de' size='50' type='text' name='"+e+"pn' value='"+e+"'>",newline+="<span style='font-size:12px;font-weight:bold;'> Qty: </span><input style='font-size:10px;' name='part*"+t+"*qty' type='text' size='4' value='"+n+"'>",newline+="<span style='font-size:12px;font-weight:bold;'> Price: </span><input size='10' style='font-size:10px' name='part*"+t+"*pr' type='text' value='"+l+"'>",newline+=" <input style='font-size:10px' name='part*"+t+"*button' type='button' value='Remove' onclick='deletePartItem(\""+t+"\")'><br id='part*"+t+"*br'>",document.getElementById("estimator").style.top="19%",document.getElementById("estimator").style.height="80%",document.getElementById("pestimatorqueue").style.display="block",document.getElementById("pestimatorqueue").innerHTML+=newline,document.getElementById("estimatorqueue").style.display="block",document.guideform.partitems.value+=t+"~"}function removeSpaces(e){return e.split(" ").join("")}function deletePartItem(e){nelement=document.getElementById("part*"+e+"*pn"),velement=document.getElementById("part*"+e+"*de"),belement=document.getElementById("part*"+e+"*button"),brelement=document.getElementById("part*"+e+"*br"),qelement=document.getElementById("part*"+e+"*qty"),pelement=document.getElementById("part*"+e+"*pr"),nelement.style.display="none",velement.style.display="none",belement.style.display="none",brelement.style.display="none",qelement.style.display="none",pelement.style.display="none",nelement.parentNode.removeChild(nelement),velement.parentNode.removeChild(velement),belement.parentNode.removeChild(belement),brelement.parentNode.removeChild(brelement),qelement.parentNode.removeChild(qelement),pelement.parentNode.removeChild(pelement),document.guideform.partitems.value.replace(e+"~","")}function deleteLaborItem(e){delement=document.getElementById("labor*"+e+"*de"),nelement=document.getElementById("labor*"+e+"*ti"),velement=document.getElementById("labor*"+e+"*tech"),belement=document.getElementById("labor*"+e+"*button"),brelement=document.getElementById("labor*"+e+"br"),delement.style.display="none",nelement.style.display="none",velement.style.display="none",belement.style.display="none",brelement.style.display="none",delement.parentNode.removeChild(delement),nelement.parentNode.removeChild(nelement),velement.parentNode.removeChild(velement),belement.parentNode.removeChild(belement),brelement.parentNode.removeChild(brelement),document.guideform.laboritems.value.replace(e+"~","")}function showPayment(e){document.getElementById("popup").style.display="block",document.getElementById("popuphider").style.display="block","swipe"==e&&(document.pmtform.processtype.value="cc",document.getElementById("cctable").style.display="",document.getElementById("achtable").style.display="none",document.getElementById("othertable").style.display="none",focusPayment()),"echeck"==e&&(document.pmtform.processtype.value="ach",document.getElementById("cctable").style.display="none",document.getElementById("achtable").style.display="",document.getElementById("othertable").style.display="none"),"cash"==e&&(document.pmtform.processtype.value="cash",document.getElementById("cctable").style.display="none",document.getElementById("achtable").style.display="none",document.getElementById("othertable").style.display="")}function closePayment(){document.getElementById("popup").style.display="none",document.getElementById("popuphider").style.display="none"}
<%
if merchantaccount = "no" then
%>
function savePayment(){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	var url="ropostpayment.asp?roid="+document.pmtform.roid.value+"&amt="+document.pmtform.amt.value+"&type="+document.pmtform.type.value+"&cid="+document.pmtform.cid.value+"&number="+document.pmtform.number.value+"&shopid=<%=shopid%>"
	xmlHttp.onreadystatechange=paymentStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}
<%
else
%>
function savePayment(){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	var url="ropostpayment.asp?roid="+document.pmtform.roid.value+"&amt="+document.pmtform.cashamt.value+"&type="+document.pmtform.cashtype.value+"&cid="+document.pmtform.cid.value+"&number="+document.pmtform.cashnumber.value+"&shopid=<%=shopid%>"
	//alert(url)
	xmlHttp.onreadystatechange=paymentStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}


<%
end if
%>
function isNumber(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function postCCPayment(){
	//alert(document.pmtform.processtype.value)
	if(document.pmtform.processtype.value == "cc"){
		var cardswipe = document.pmtform.cardswipe.value
	    echar = cardswipe.indexOf("E?")
	    if(echar == -1 && cardswipe.length >= 15){
	    	
	    	// set the cardstring field to no
	    	document.pmtform.cardstring.value = 'yes'
	    	if(!isNumber(document.pmtform.ccamt.value)){
	    		alert("You must enter the amount you wish to process")
	    		return
	    	}
	    }else{
	    
	     	// set the cardstring field to yes
	    	document.pmtform.cardstring.value = 'no'
	    	//alert(document.pmtform.cardstring.value)
	   
	    	 //process with the swipe string only
	    	if(!isNumber(document.pmtform.cc.value)){
	    		alert("You must enter the amount you wish to process")
	    		return
	    	}
	    	if (document.pmtform.ccfirstname.value.length == 0){
	    		alert("You must enter the Card Holder FIRST name")
	    		return
	    	}
	    	if (document.pmtform.cclastname.value.length == 0){
	    		alert("You must enter the Card Holder LAST name")
	    		return
	    	}
	
	    	
	    }
	    
		//complete the process with cardstring only
		xmlHttp=GetXmlHttpObject()
		if (xmlHttp==null){
			alert ("Browser does not support HTTP Request")
			return
		} 
		cardswipe = cardswipe.replace(/\%/,"@")
		document.getElementById("cctimer").style.display="block"
		document.getElementById("hider").style.display="block"
		ptype = document.pmtform.processtype.value
		//alert(document.pmtform.cardstring.value)
		var url="ropostcc.asp?processtype="+ptype+"&expmonth="+document.pmtform.expmonth.value+"&expyear="+document.pmtform.expyear.value+"&cc="+document.pmtform.cc.value+"&ccfirstname="+document.pmtform.ccfirstname.value+"&cclastname="+document.pmtform.cclastname.value+"&cardstring="+document.pmtform.cardstring.value+"&cardswipe="+cardswipe+"&roid="+document.pmtform.roid.value+"&amt="+document.pmtform.ccamt.value+"&type="+document.pmtform.cctype.value+"&cid="+document.pmtform.cid.value+"&number="+document.pmtform.ccnumber.value+"&shopid=<%=shopid%>"
		//alert(url)
		xmlHttp.onreadystatechange=postccstatechanged 
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
	}else{
	    

    	 //process with the swipe string only
    	if(!isNumber(document.pmtform.accountnumber.value)){
    		alert("You must enter the amount you wish to process")
    		return
    	}
    	if (document.pmtform.achfirstname.value.length == 0){
    		alert("You must enter the Card Holder FIRST name")
    		return
    	}
    	if (document.pmtform.achlastname.value.length == 0){
    		alert("You must enter the Card Holder LAST name")
    		return
    	}
    	if (document.pmtform.routing.value.length == 0){
    		alert("You must enter the Card Holder LAST name")
    		return
    	}

		//complete the process with cardstring only
		xmlHttp=GetXmlHttpObject()
		if (xmlHttp==null){
			alert ("Browser does not support HTTP Request")
			return
		} 

		document.getElementById("cctimer").style.display="block"
		document.getElementById("hider").style.display="block"
		ptype = document.pmtform.processtype.value
		//alert(document.pmtform.cardstring.value)
		var url="ropostcc.asp?routing="+document.pmtform.routing.value+"&accountnumber="+document.pmtform.accountnumber.value+"&processtype="+ptype+"&achfirstname="+document.pmtform.achfirstname.value+"&achlastname="+document.pmtform.achlastname.value+"&roid="+document.pmtform.roid.value+"&achamt="+document.pmtform.achamt.value+"&shopid=<%=shopid%>"
		//alert(url)
		xmlHttp.onreadystatechange=postccstatechanged 
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
	}
}

function postccstatechanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,"approved"==rt?(alert("Transaction approved and payment posted to Repair Order"),document.getElementById("cctimer").style.display="none",document.getElementById("hider").style.display="none",document.getElementById("popup").style.display="none",document.getElementById("popuphider").style.display="none",document.body.onunload="",saveAll("ro.asp?a=n&","","","","","")):(alert(rt),document.getElementById("cctimer").style.display="none",document.getElementById("hider").style.display="none"))}function paymentStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,"posted"==rt?(document.getElementById("popup").style.display="none",document.getElementById("popuphider").style.display="none",document.body.onunload="",saveAll("ro.asp?","","","","")):alert(rt))}
function deletePayment(id){

	c = confirm("This will delete this payment.  Are you sure?")
	if(c){
		document.body.onunload = ""
		saveAll("rodeletepayment.asp?roid=<%=request("roid")%>&id="+id+"&",'','','','')
	}

}

function clearPercentFee(e){document.getElementById(e).value="zero"}function showImage(e){alert(e),window.open(e,"image")}function clearFee(n){c=confirm("This will set this fees percentage to 0.  This cannot be reversed on this RO.  Are you sure?"),c&&(feetoclear=eval("document.theform.UserFee"+n),percenttoclear=eval("document.theform.dbuserfee"+n+"percent"),dbfeetoclear=eval("document.theform.dbUserFee"+n),feetoclear.value="0",percenttoclear.value="0",dbfeetoclear.value="0",calcPercents())}function showRecRepairs(){document.getElementById("recrepairs").style.display="block",document.getElementById("popuphider").style.display="block"}function closeRecRepairs(){document.getElementById("recrepairs").style.display="none",document.getElementById("popuphider").style.display="none"}function showVehicleFields(){document.getElementById("customvehiclefields").style.display="block",document.getElementById("popuphider").style.display="block"}function hideVehicleFields(){document.getElementById("customvehiclefields").style.display="none",document.getElementById("popuphider").style.display="none"}function showSpouse(){document.getElementById("spouseinfo").style.display="block",document.getElementById("popuphider").style.display="block"}function hideSpouse(){document.getElementById("spouseinfo").style.display="none",document.getElementById("popuphider").style.display="none"}function showDisclosure(){document.getElementById("disclosure").style.display="block",document.getElementById("popuphider").style.display="block"}function hideDisclosure(){document.getElementById("disclosure").style.display="none",document.getElementById("popuphider").style.display="none"}
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
				<%
				if lcase(chargeshopfeeson) = "all" then
				%>
				fee1val = (document.theform.dbuserfee1percent.value * document.theform.dbSubtotal.value)/100
				<%
				elseif lcase(chargeshopfeeson) = "labor" then
				%>
				fee1val = (document.theform.dbuserfee1percent.value * document.theform.dbTotalLbr.value)/100
				<%
				end if
				%>
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
				<%
				if request.cookies("shopid") <> "1734" then
				%>
				fee2val = (document.theform.dbuserfee2percent.value * document.theform.dbSubtotal.value)/100
				<%
				else
				%>
				fee2val = (document.theform.dbuserfee2percent.value * document.theform.dbTotalPrts.value)/100
				<%
				end if
				%>
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
				<%
				if lcase(chargeshopfeeson) = "all" then
				%>
				fee3val = (document.theform.dbuserfee3percent.value * document.theform.dbSubtotal.value)/100
				<%
				elseif lcase(chargeshopfeeson) = "labor" then
				%>
				fee3val = (document.theform.dbuserfee3percent.value * document.theform.dbTotalLbr.value)/100
				<%
				end if
				%>

				if(userfee3max > 0 && userfee3max < fee1val){fee1val = userfee3max}
				document.theform.UserFee3.value = fee3val
				document.theform.dbUserFee3.value = fee3val
				fee3 = fee3val
			}
		}
	}

	if(fee1.length == 0){
		fee1 = 0
	}
	if(fee2.length == 0){
		fee2 = 0
	}
	if(fee3.length == 0){
		fee3 = 0
	}

	hw = document.theform.HazardousWaste.value
	sf = document.theform.storagefee.value
	document.theform.dbTotalFees.value = parseFloat(fee1)+parseFloat(fee2)+parseFloat(fee3)+parseFloat(hw)+parseFloat(sf)
}

function declineAcceptIssue(e){saveAll("roissuedeclineaccept.asp?comid="+e+"&")}function hideDrops(e){document.getElementById(e).style.display="none"}function deleteRec(e){if(xmlHttp=GetXmlHttpObject(),null==xmlHttp)return void alert("Browser does not support HTTP Request");vin=document.theform.dbVin.value,r=Math.floor(111111*Math.random());var t="rodeleterec.asp?id="+e;xmlHttp.onreadystatechange=recChanged,xmlHttp.open("GET",t,!0),xmlHttp.send(null)}
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

function addRecRepair(){document.getElementById("recrepairs").style.display="none",document.getElementById("addrecrepairs").style.display="block"}function cancelAddRec(){document.getElementById("addrecrepairs").style.display="none",document.getElementById("popuphider").style.display="none"}function addRecRep(e){if(xmlHttp=GetXmlHttpObject(),null==xmlHttp)return void alert("Browser does not support HTTP Request");qs="&desc="+document.getElementById("desc").value+"&totalrec="+document.getElementById("recamt").value,vin=document.theform.dbVin.value,r=Math.floor(111111*Math.random());var t="addrecrep.asp?roid="+e+qs;xmlHttp.onreadystatechange=addRecChanged,xmlHttp.open("GET",t,!0),xmlHttp.send(null)}function addRecChanged(){
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
function editRecRep(){if(xmlHttp=GetXmlHttpObject(),null==xmlHttp)return void alert("Browser does not support HTTP Request");qs="?id="+document.getElementById("editrecid").value+"&desc="+document.getElementById("editdesc").value+"&totalrec="+document.getElementById("editrecamt").value,vin=document.theform.dbVin.value,r=Math.floor(111111*Math.random());var e="editrecrep.asp"+qs;xmlHttp.onreadystatechange=editRecChanged,xmlHttp.open("GET",e,!0),xmlHttp.send(null)}function editRecChanged(){
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
function editRec(e,t,d){document.getElementById("recrepairs").style.display="none",document.getElementById("editrecrepairs").style.display="block",document.getElementById("editrecamt").value=d,document.getElementById("editrecid").value=e,document.getElementById("editdesc").value=t}function cancelEditRec(){document.getElementById("editrecrepairs").style.display="none",document.getElementById("popuphider").style.display="none"}function declineItem(e){c=confirm("This will decline the repairs for this issue AND move all parts and labor to the Recommended Repairs.  Are you sure?"),c&&(document.body.onunload="",document.theform.action="1238savero.asp?comid="+e+"&comstat=Declined",saveAll("ro.asp?a=n&","","","","",""))}
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

function comStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&("error"!=xmlHttp.responseText?(document.body.onunload="",saveAll("ro.asp?a=n&","","","","","")):alert(xmlHttp.responseText))}
function timeClock(laborid,comid,tech,labortimeclockid){

	//alert(getLocalDate())
	//post time to labor time clock
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	if(labortimeclockid == "x"){
		//document.getElementById("clock"+laborid).style.backgroundColor = "white"
		//document.getElementById("clock"+laborid).innerHTML = '<img alt="" src="newimages/runningman.gif" height="17" width="17">Clock Out'
		var url="rolabortimeclock.asp?roid=<%=request("roid")%>&comid="+comid+"&tech="+tech+"&laborid="+laborid+"&labortimeclockid="+labortimeclockid
		r=Math.floor(Math.random()*111111)
		url += "&r=" + r
		currTime = getLocalDate()
		url += "&currtime=" + currTime
		xmlHttp.onreadystatechange=timeClockStateChanged 
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
	}else{
		document.getElementById("clock"+laborid).style.backgroundColor = "yellow"
		document.getElementById("clock"+laborid).innerHTML = '<img alt="" src="newimages/clock.png">Clock In'
		
		// show tech window
		ed = getLocalDate()
		getTechForTime(tech,comid,laborid,ed,labortimeclockid)
	}
}

function timeClockStateChanged(){if(4==xmlHttp.readyState||"complete"==xmlHttp.readyState){var e=xmlHttp.responseText;if("success"==e)return document.body.onunload="",void saveAll("ro.asp?a=n&","","","","","");if("error"==e)return alert("There was an error.  Please click Calculate, then try again"),document.getElementById("clock"+laborid).style.backgroundColor="yellow",void(document.getElementById("clock"+laborid).innerHTML='<img alt="" src="newimages/clock.png">Clock In');if(e.indexOf("*")>0){var t=e.split("*");return alert("Please clockout on RO "+t[1]+" before clocking in on this RO"),document.getElementById("clock"+laborid).style.backgroundColor="yellow",void(document.getElementById("clock"+laborid).innerHTML='<img alt="" src="newimages/clock.png">Clock In')}}}function getTechForTime(e,t,o,l,n){document.getElementById("popuphider").style.display="block",document.getElementById("techfortime").style.display="block",document.techfortimeform.comid.value=t,document.techfortimeform.laborid.value=o,document.techfortimeform.enddate.value=l,document.techfortimeform.labortimeclockid.value=n;for(var c=e,r=document.techfortimeform.techformtimetech,i=0;i<r.options.length;i++)if(r.options[i].text==c){r.selectedIndex=i;break}}function getLocalDate(){var e=new Date,t=e.getMonth()+1,o=e.getDate(),l=e.getFullYear(),e=new Date,n=e.getHours(),c=e.getMinutes(),r=e.getSeconds();return t+"/"+o+"/"+l+" "+n+":"+c+":"+r}function getCookie(e){var t,o,l,n=document.cookie.split(";");for(t=0;t<n.length;t++)if(o=n[t].substr(0,n[t].indexOf("=")),l=n[t].substr(n[t].indexOf("=")+1),o=o.replace(/^\s+|\s+$/g,""),o==e)return unescape(l)}function setCookie(e,t,o){var l=new Date;l.setDate(l.getDate()+o);var n=escape(t)+(null==o?"":"; expires="+l.toUTCString());document.cookie=e+"="+n}function postLaborTimeClock(){if(xmlHttp=GetXmlHttpObject(),null==xmlHttp)return void alert("Browser does not support HTTP Request");var e="rolabortimeclock.asp?tech="+document.techfortimeform.techformtimetech.value+"&enddate="+document.techfortimeform.enddate.value+"&labortimeclockid="+document.techfortimeform.labortimeclockid.value;xmlHttp.onreadystatechange=laborTimeStateChanged,xmlHttp.open("GET",e,!0),xmlHttp.send(null)}function cancelTimeClock(){laborid=document.getElementById("timeclocklaborid").value,document.getElementById("techfortime"+laborid).style.display="none",document.getElementById("popuphider").style.display="none",document.getElementById("popuphider").style.display="none"}function laborTimeStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(document.getElementById("popuphider").style.display="none",document.getElementById("techfortime").style.display="none",document.body.onunload="",saveAll("ro.asp?a=n&"))}function showClockList(e,t){document.getElementById("hider").style.display="block",document.getElementById("clocklist2-"+e).style.position="absolute",curtopleft=findMouse(t).toString(),tar=curtopleft.split(","),document.getElementById("clocklist2-"+e).style.top=tar[1],document.getElementById("clocklist2-"+e).style.left=tar[0],document.getElementById("clocklist2-"+e).style.display="block"}function closeClockList(e){document.getElementById("hider").style.display="none",document.getElementById("clocklist2-"+e).style.display="none"}function findMouse(e){for(var t=curtop=0,o=e,l=!1;(o=o.parentNode)&&o!=document.body;)t-=o.scrollLeft||0,curtop-=o.scrollTop||0,"fixed"==getStyle(o,"position")&&(l=!0);if(l&&!window.opera){var n=scrollDist();t+=n[0],curtop+=n[1]}do t+=e.offsetLeft,curtop+=e.offsetTop;while(e=e.offsetParent);return[t,curtop]}function showParts(e){window.open(e,"po","toolbar=no,menubar=no,scrollbars=yes,resizable=yes")}function deleteVehIssue(e){c=confirm("This will delete this issue and all parts and labor associated with it.  Are you sure?"),c&&saveAll("deletecomplaintfromro.asp?complaintid="+e+"&","","","","","")}function setReminder(){document.getElementById("popuphider").style.display="block",document.getElementById("reminders").style.display="block"}function deleteReminder(e){c=confirm("This will delete this reminder.  Are you sure?"),c&&processDeleteReminder(e)}<%
if carfax <> "no" then
%>
function showCFHistory(vin,lic,state){
	document.getElementById("carfax").src = ""
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="carfax/vehiclehistory.asp?vin="+vin+"&lic="+lic+"&state="+state+"&shopid=<%=request.cookies("shopid")%>"
	//alert(url)
	xmlHttp.onreadystatechange=cfHistoryStateChanged
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
	
}

function cfHistoryStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,x=rt.indexOf("https"),x>=0?(document.getElementById("popuphider").style.display="block",document.getElementById("carfaxbutton").style.display="block",document.getElementById("carfax").style.display="block",document.getElementById("carfax").src=rt):alert(rt))}function closeCarfax(){document.getElementById("popuphider").style.display="none",document.getElementById("carfax").style.display="none",document.getElementById("carfaxbutton").style.display="none"}<%
else
%>
function showCFHistory(){

	alert("Carfax history is not available until you register your shop with Carfax.  Please click the Carfax icon on the WIP screen")
	
}

<%
end if
%>
function processDeleteReminder(id){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="rodeletereminder.asp?id="+id+"&shopid=<%=shopid%>"
	//alert(url)
	xmlHttp.onreadystatechange=processDeleteReminderStateChanged
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function processDeleteReminderStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,alert("success"==rt?"Reminder Deleted":rt),document.getElementById("popuphider").style.display="none",document.getElementById("reminders").style.display="none")}function closeReminder(){document.getElementById("popuphider").style.display="none",document.getElementById("reminders").style.display="none"}function addReminder(){document.getElementById("reminders").style.display="none",document.getElementById("popuphider").style.display="block",document.getElementById("addreminder").style.display="block"}function closeAddReminder(){document.getElementById("popuphider").style.display="none",document.getElementById("addreminder").style.display="none"}
function saveReminder(){

	if (document.remform.reminder.value.length == 0){
		alert("Please enter the reason for the reminder")
		return
	}
	if (document.remform.reminderdate.value.length == 0){
		alert("Please enter the date for the reminder")
		return
	}
	if (document.remform.recurring.value == "Select"){
		alert("Please select a value for Recurring")
		return
	}
	if (document.remform.recurring.value == "yes"){
		if(document.remform.interval.value.length == 0){
			alert("You have selected Recurring.  Please enter the number of days between reminders in the Interval box")
			return
		}else{
			if (!isNumber(document.remform.interval.value)){
				alert("Please enter only a number in the Interval box")
				return
			}
		}
	}

	reminder = document.remform.reminder.value
	reminderdate = document.remform.reminderdate.value
	recurring = document.remform.recurring.value
	recurringinterval = document.remform.interval.value
	
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="roaddreminder.asp?shopid=<%=shopid%>&customerid=<%=cid%>&vehid=<%=vehid%>&reminder="+reminder+"&reminderdate="+reminderdate+"&recurring="+recurring+"&interval="+recurringinterval
	//alert(url)
	xmlHttp.onreadystatechange=saveReminderStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function saveReminderStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,alert("success"==rt?"Reminder has been added":rt),document.getElementById("addreminder").style.display="none",document.getElementById("popuphider").style.display="none")}function editDate(){document.getElementById("datebox").style.display="block"}function printDiv(e){document.getElementById("b1").style.display="none",document.getElementById("b2").style.display="none",document.getElementById("b3").style.display="none";var t=document.getElementById(e).innerHTML,n=document.body.innerHTML;document.body.innerHTML=t,window.print(),document.body.innerHTML=n,document.getElementById("b1").style.display="block",document.getElementById("b2").style.display="block",document.getElementById("b3").style.display="block"}function showEmailAddress(e){document.getElementById("invoicediv").style.display="none",document.getElementById("invoice").style.display="none",document.getElementById("hider").style.display="none",document.getElementById("sendemail").style.display="block",document.getElementById("hider").style.display="block",document.getElementById("inclprint").value=e,document.getElementById("printdrop").style.display="none"}
<%
if esign = "yes" then
%>
function redirectWithEmail(){

	var invmessage = document.getElementById("emailinvoicemessage").value
	invmessage = invmessage.replace(new RegExp('\r?\n','g'), '|');

	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("invoice").style.display = "none"
	//document.getElementById("hider").style.display = "none"
	var inclprint = document.getElementById("inclprint").value
	var curremail = document.getElementById("curremailaddress").value
	if(document.getElementById("updateemail").checked ==true){
		updateemail = "yes"
	}
	if(document.getElementById("updateemail").checked ==false){
		var updateemail = "no"
	}
	
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	email = document.getElementById("curremailaddress").value
	if (document.getElementById("updateemail").checked ==true){update = "yes"}else{update="no"}
	var pdfpath = document.getElementById("invoice").src
	//alert(pdfpath)
	var url="roemailinvoice.asp?email="+email+"&update="+update+"&pdfpath="+pdfpath+"&shopid=<%=shopid%>&message="+invmessage
	//alert(url)
	xmlHttp.onreadystatechange=function(){
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
				rt = xmlHttp.responseText
				if (rt == "success"){
					alert("Email sent.")
					document.getElementById("hider").style.display = "none"
					document.getElementById("sendemail").style.display = "none"
				}else{
					alert(rt)
				}
			}
		}

	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}
<%
else
%>
function redirectWithEmail(){

	var inclprint = document.getElementById("inclprint").value
	var curremail = document.getElementById("curremailaddress").value
	if(document.getElementById("updateemail").checked ==true){
		updateemail = "yes"
	}
	if(document.getElementById("updateemail").checked ==false){
		var updateemail = "no"
	}
	if(inclprint == "yes"){
		document.body.onunload = ''
		location.href='ro.asp?updateemail='+updateemail+'&roid=<%=request("roid")%>&emro=yes&printro=yes&email='+curremail
	}
	if(inclprint == "no"){
		document.body.onunload = ''
		location.href='ro.asp?updateemail='+updateemail+'&roid=<%=request("roid")%>&emro=yes&printro=no&email='+curremail
	}


}
<%
end if
%>
function findPosPrint(obj) {

  if (document.getElementById("requiremileagetoprintro").value == "Yes"){
  	if (document.theform.MilesIn.value.length == 0){
  		alert("You must enter vehicle mileage in to print a repair order")
  		return
  	}
  }
  	
  if (document.getElementById("printdrop").style.display == "inline-block"){
  	document.getElementById("printdrop").style.display = "none"
  	return
  }
  <%
  response.write "//" & lcase(requiretirepressuretoprintro)
  if lcase(requiretirepressuretoprintro) = "yes" then
  %>
  tpinlf = document.getElementById("tirepressureinlf").value.length
  tpinlr = document.getElementById("tirepressureinlr").value.length
  tpinrf = document.getElementById("tirepressureinrf").value.length
  tpinrr = document.getElementById("tirepressureinrr").value.length
  
  tpoutlf = document.getElementById("tirepressureoutlf").value.length
  tpoutlr = document.getElementById("tirepressureoutlr").value.length
  tpoutrf = document.getElementById("tirepressureoutrf").value.length
  tpoutrr = document.getElementById("tirepressureoutrr").value.length
  
  tdlf = document.getElementById("treaddepthlf").value.length
  tdlr = document.getElementById("treaddepthlr").value.length
  tdrf = document.getElementById("treaddepthrf").value.length
  tdrr = document.getElementById("treaddepthrr").value.length
  
  if(document.theform.Status.value == 'FINAL' || document.theform.Status.value == 'CLOSED'){
  	if (tpinlf == 0 || tpinlr == 0 || tpinrf == 0 || tpinrr == 0 || tpoutlf == 0 || tpoutlr == 0 || tpoutrf == 0 || tpoutrr == 0 || tdlf == 0 || tdlr == 0 || tdrf == 0 || tdrr == 0){
	  	alert("You must enter all tire pressures and tread depths to print a Final Invoice")
	  	return
	}
  }
  <%
  end if
  %>
  
  // now show the menu

  //document.getElementById("printdrop").style.left = "45%"
 // document.getElementById("printdrop").style.top = "10%"
  document.getElementById("printdrop").style.display = 'inline-block'
  document.getElementById("hider").style.display = "block"

}

function showInspection(){

	var r=Math.floor(Math.random()*100001)
	<%
	' check for DVI
	dvistmt = "select * from companyadds where shopid = '" & shopid & "' and `name` = 'DVI Boss'"
	set dvirs = con.execute(dvistmt)
	if not dvirs.eof then
	%>
	document.getElementById("inspection").src = "inspection/dvi.asp?roid=<%=request("roid")%>&r="+r 
	<%
	else
	%>
	document.getElementById("inspection").src = "inspection/inspection-new.asp?roid=<%=request("roid")%>&r="+r 
	<%
	end if
	%>

	document.getElementById("inspection").style.display = "block"
	document.getElementById("popuphider").style.display = "block"
	document.getElementById("inspectionbutton").style.display = "block"

}

function closeInspection(){

	document.getElementById("inspection").src = "" 
	document.getElementById("inspection").style.display = "none"
	document.getElementById("inspectionbutton").style.display = "none"
	document.getElementById("popuphider").style.display = "none"


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

function postAllParts(){

	document.getElementById("hider").style.display = "block"
	document.getElementById("timer").style.display = "block"
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="ropostallparts.asp?roid=<%=request("roid")%>&shopid=<%=shopid%>"
	//alert(url)
	xmlHttp.onreadystatechange=postPartsStateChanged
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function postPartsStateChanged(){(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,alert("success"==rt?"All parts posted to accounting":rt),document.getElementById("hider").style.display="none",document.getElementById("timer").style.display="none")}
function reprintRO(n){

	filepath = "savedinvoices/<%=shopid%>/<%=request("roid")%>/"+n
	
	document.getElementById("invoice").src = filepath
	document.getElementById("invoice").style.width = "100%"
	document.getElementById("invoice").style.display = "block"
	document.getElementById("invoicediv").style.display = "block"
	document.getElementById('invoice').reload(true)
}
function isIpad(){null!=navigator.userAgent.match(/iPad/i)||document.getElementById("printbutton")&&(document.getElementById("printbutton").style.display="")}
function printROwithSave(){
	document.getElementById('printdrop').style.display='none';
	document.getElementById('hider').style.display='none';
	document.body.onunload='';
	if(navigator.userAgent.match(/iPad/i) != null){
		location.href='ro.asp?roid=<%=request("roid")%>&printro=yes';
	}else{
		saveAll('ro.asp?printro=yes&','','','','','')
	}
}

function loadGP(){
	document.getElementById("gpdetails").src='gp.asp?discount='+document.getElementById("discountforgp").value+'&shopid=<%=shopid%>&roid=<%=request("roid")%>&subtotal='+document.getElementById("totalwithouttax").value
	document.getElementById("gp").style.display = "block"

}

function showGPDetails(){

	document.getElementById('gpdetails').style.display='block'
	document.getElementById('hider').style.display='block'

}
function completeItem(type,id,stat){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="rochangepartslaborstatus.asp?id="+id+"&status="+stat+"&type="+type+"&shopid=<%=request.cookies("shopid")%>"
		
	//alert(url)
	xmlHttp.onreadystatechange=function(){
									if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
										//alert(xmlHttp.responseText)
										tempxmlhttp = GetXmlHttpObject()
										turl = "complaintlayout.asp?roid=<%=request("roid")%>&shopid=<%=request.cookies("shopid")%>"
										
										tempxmlhttp.onreadystatechange = function(){
																				if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
																					document.getElementById("issues").innerHTML = tempxmlhttp.responseText
																				}
																			}
										tempxmlhttp.open("GET",turl,true)
										tempxmlhttp.send(null)
									}
								}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

<%
if showgp = "yes" then
%>
setTimeout("loadGP()",1000)
<%
end if
%>

function addToRO(recid,comid){

	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="roaddoldrecommendtoro.asp?recid="+recid+"&comid="+comid+"&shoid=<%=request.cookies("shopid")%>&roid=<%=request("roid")%>"
		
	//alert(url)
	xmlHttp.onreadystatechange=function(){
		if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
			rt = xmlHttp.responseText
			if(rt == "success"){
				// calculate the ro
				document.theform.Status.focus();
				document.getElementById("recrepairs").style.display = "none"
				document.body.onunload='';
				saveAll('ro.asp?a=n&','','','','','')
			}else{
				alert(rt)
			}
		}
	}
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)


}

function restoreAllRecs(){

	// get all selected checkboxes
	rrids = ""
	cbs = document.getElementsByTagName("input")
	console.log(cbs.length)
	for (j=0;j<cbs.length;j++){
		if(cbs[j].type == "checkbox" && document.getElementById(cbs[j].id).checked && cbs[j].id.substring(0,4) == "rrcb"){
			rrids = rrids + cbs[j].id.replace("rrcb","") + ","
		}
	}
	rrids = rrids.substring(0,rrids.length-1)
	if(rrids.length < 2){
		alert("You must check off the Recommendations you wish to add")
		return
	}
	url = "roaddallrecs.asp?roid=<%=request("roid")%>&shopid=<%=request.cookies("shopid")%>&recids="+rrids
	console.log(url)
	c = confirm("This will restore all checked recommendations.  Are you sure?")
	
	if(c){
		xmlHttp=GetXmlHttpObject()
		xmlHttp.onreadystatechange=function(){
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
				rt = xmlHttp.responseText
				if(rt == "success"){
					// calculate the ro
					document.theform.Status.focus();
					document.getElementById("recrepairs").style.display = "none"
					document.body.onunload='';
					saveAll('ro.asp?a=n&','','','','','')
				}else{
					console.log(rt)
				}
			}
		}
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
	}

}

function openTPMS(){
	<%
	' get tpms account if it exists
	stmt = "select * from tpmsaccounts where shopid = '" & request.cookies("shopid") & "'"
	set tprs = con.execute(stmt)
	if not tprs.eof then
	%>
	URL = "http://portal.tpmsmanager.com?u=<%=tprs("username")%>&pwd=<%=tprs("password")%>&vin=" + document.getElementById("thevinnumber").innerHTML
	window.open(URL)
	<%
	else
	%>
	alert("You do not have a TPMS Speed Account.  If you would like to add it to your subscription, please click the TPMSSpeed icon on the WIP screen")
	<%
	end if
	%>
}

function launchNexpart(){
	
	document.getElementById("nexpartsupplier").style.display = "none"
	document.getElementById("popuphider").style.display = "none"
	document.getElementById("hider").style.display = "none"

	npcreds = document.getElementById("nexpartselect").value
	console.log("npcreds:"+npcreds)
	if (npcreds.indexOf("|") >= 0){
		npar = npcreds.split("|")
		npuser = npar[0]
		nppwd = npar[1]
		if (npuser.length > 0 && nppwd.length > 0){
			h = screen.availHeight - 80
			w = screen.availWidth - 20
			URL = 'http://www.nexpart.com/extpart.php?sms_form=X&provider=NEXLINKCSBTECH&pwd='+nppwd+'&nexpartuname='+npuser+'&vin=<%=rs("vin")%>&identifier=<%=request("shopid")%>-<%=request("roid")%>&webpost=https://<%=request.servervariables("SERVER_NAME")%>/sbp/nexpart/return.asp'
			str='addressbar=0,toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=yes,screenX=0,screenY=0,top=0,left=0,maximize=1,'
			str=str+'height='+h+', width='+w
			//eval('var ' + id+'=window.open(URL, id,str )')
			var mywin = window.open(URL, "npwin",str )
		}else{
			alert("You cannot order parts through Nexpart without a valid username and password.  If you have one, go to Settings -> Integrations and enter them there")
		}
	}else{
		alert("You must have and select a valid Nexpart Account")
	}
}

function launchWorldpac(){

<%if worldpac = "yes" then%>
	URL = 'http://<%=request.servervariables("SERVER_NAME")%>/sbp/worldpac/main.asp?shopid=<%=shopid%>&roid=<%=request("roid")%>'
	var mywin = window.open(URL, "npwin" )
	myx = setInterval(function(){
		console.log("checking closed")
		if (mywin.closed){
			window.location.reload()
			clearInterval(myx)
		}
	},1000)
<%else%>
	alert("You must add WorldPac parts ordering in settings before you can use this integration")
<%end if%>
}

function showEpicor(){
	
	document.getElementById("hider").style.display = "block"
	document.getElementById("epicor").style.display = "block"
	document.getElementById("orderselect").style.display = "none"
	//winPopUp("epicor/main.asp?roid=<%=request("roid")%>")

}

function showOrderSelect(){
	document.getElementById("hider").style.display = "block"
	document.getElementById("orderselect").style.display = "block"

}

function closeOrderingSelect(){
	document.getElementById("hider").style.display = "none"
	document.getElementById("orderselect").style.display = "none"

}
function hideEpicor(){
	
	document.getElementById("hider").style.display = "none"
	document.getElementById("epicor").style.display = "none"
	//winPopUp("epicor/main.asp?roid=<%=request("roid")%>")

}

function winPopUp(URL) {
	if(document.getElementById('epicorselect').value != "select"){
		day = new Date()
		id = day.getTime();
		h = screen.availHeight - 80
		w = screen.availWidth - 20
		str='addressbar=0,toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=yes,screenX=0,screenY=0,top=0,left=0,maximize=1,'
		str=str+'height='+h+', width='+w
		eval('page' + id+'=window.open(URL, id,str )')
		hideEpicor()
	}else{
		alert("Please select a supplier to login to")
	}
}

function launchEpicor(){

	if (document.getElementById("epicorselect").value != 'select' && document.getElementById("om").value != 'select'){
		om = document.getElementById("om").value
		winPopUp('epicor/main.asp?om='+om+'&vin=<%=rs("vin")%>&supplier='+document.getElementById('epicorselect').value+'&roid=<%=request("roid")%>')
	}else{
		alert("You must enter your Epicor Parts Ordering Credentials in Settings -> Integrations")
	}

}

function changeVIOrder(){

	$('#issueorder').show()
	$('#issueorder').attr("src","complaintaddupdatedd.asp?roid=<%=request("roid")%>") 
	$('#popuphider').show()

}

function hideVIOrder(){
	$('#issueorder').hide()
	$('#issueorder').attr("src","") 
	$('#popuphider').hide()

}

function showlaunchNexpart(){

	document.getElementById("orderselect").style.display = "none"
	document.getElementById("nexpartsupplier").style.display = "block"
	document.getElementById("popuphider").style.display = "block"
	

}

</script>
</head>
<%
if request("showrec") = "y" then srr = "showRecRepairs()"
%>

<body link="#800000" vlink="#800000" alink="#800000" onload="isIpad();ttlFees();closeTimer();<%=srr%>"  onunload="" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<div style="display:none;position:absolute;left:30%;top:100px;border:2px black solid;background-color:white;z-index:2500;width:40%;padding:10px;text-align:center" id="nexpartsupplier">
<b>NOTE:</b> make sure your pop up blocker is OFF.  <span style="color:red;font-weight:bold">Select a supplier below, then click Login.</span>
<br><br>
<%
%>
<select id="nexpartselect">
	<option value="select">Select Supplier</option>
	<%
	stmt = "select * from nexpart where shopid = '" & shopid & "'"
	set nexrs = con.execute(stmt)
	if not nexrs.eof then
		do until nexrs.eof
	%>
	<option value="<%=nexrs("username") & "|" & nexrs("password")%>"><%=nexrs("desc")%></option>
	<%
			nexrs.movenext
		loop
	end if
	%>
</select>
<br><br>

<input onclick="launchNexpart()" id="launchnpbutton" name="Button1" type="button" value="Login" > <input onclick="hideNexpart()" name="Button1" type="button" value="Cancel" >

</div>

<iframe style="border-radius:5px;background-color:white;z-index:1005;display:none;position:absolute;top:40px;left:10%;width:80%;height:80%;border:2px black solid" id="issueorder"></iframe>

<div style="display:none;position:absolute;left:30%;top:100px;border:2px black solid;background-color:white;z-index:2500;width:40%;padding:10px;text-align:center" id="epicor">
<b>NOTE:</b> make sure your pop up blocker is OFF.  <span style="color:red;font-weight:bold">If your username or password is incorrect, you will only see a blank window open.</span>
<br><br>
<%
stmt = "select suppliername from epicor where shopid = '" & shopid & "'"
set epicorrs = con.execute(stmt)
if not rs.eof then
%>
<select id="epicorselect">
	<option value="select">Select Supplier</option>
	<%
	do until epicorrs.eof
	%>
	<option value="<%=epicorrs("suppliername")%>"><%=epicorrs("suppliername")%></option>
	<%
		epicorrs.movenext
	loop
	%>
</select>
<%
end if
%>
<br><br>
<select id="om" name="om">
	<option value="select">Select Mode</option>
	<option value="Transfer">Build Estimate</option>
	<option value="Order">Order Parts</option>
	<option value="TransferAndOrder">Both</option>
</select>

<br><br>
<input onclick="launchEpicor()" name="Button1" type="button" value="Login" > <input onclick="hideEpicor()" name="Button1" type="button" value="Cancel" >

</div>
<div style="display:none;border-radius:10px;position:absolute;left:30%;top:100px;border:2px black solid;background-color:white;z-index:2500;width:40%;padding:10px;text-align:center" id="orderselect">
	<img onclick="closeOrderingSelect()" style="float:right;cursor:pointer" alt="" src="newimages/Close-icon.png"><br><strong>Order parts and build 
	estimates with the following providers.&nbsp; Please click the icon for your 
	preferred provider.</strong><br><br>
	<div id="epicorselectbutton" onclick="showEpicor()" style="border:1px black solid;border-radius:7px;padding:5px;width:70%;margin-left:15%;cursor:pointer"><br>
		<img alt="" height="40" src="newimages/Epicor-A3-trans.png" width="244"><br>Use for 
	Carquest, NAPA, AutoZone, Advance, etc.<br><br></div><br>
	<div id="nexpartselectbutton" onclick="showlaunchNexpart()" style="border:1px black solid;border-radius:7px;padding:5px;width:70%;margin-left:15%;cursor:pointer"><img alt="" height="66" src="../images/Powered_By_Nexpart_Logos3.png" width="250"><br>
	Use for any Nexpart powered part supplier</div><br>
	<div id="worldpacselectbutton" onclick="launchWorldpac()" style="border:1px black solid;border-radius:7px;padding:5px;width:70%;margin-left:15%;cursor:pointer">
		<img alt="" height="50" src="newimages/worldpac-trans.png" width="230"><br>Use 
	to order parts from WorldPac<br>Please select the supplier you wish to order from</div><br>
</div>

<span onclick="showGPDetails()" style="padding-left:3px;display:none;padding-top:6px;height:20px;text-align:left;font-weight:bold;position:absolute;right:5px;top:110px;font-style:normal;font-size:small;cursor:pointer;background-color:white;border:1px black solid" id="gp"></span>
<iframe style="border-radius:5px;background-color:white;z-index:1005;display:none;position:absolute;top:40px;left:20%;width:60%;height:500px;border:2px black solid" id="gpdetails"></iframe>

<div id="temp"></div>
<input type="hidden" id="pdf">
<iframe style="border-radius:5px;background-color:white;z-index:1600;display:none;position:absolute;top:40px;left:0px;width:99%;height:400px;width:700px;border:2px black solid" id="signature" src=""></iframe>
<iframe style="border-radius:5px;background-color:white;z-index:1005;display:none;position:absolute;top:40px;left:0px;width:99%;height:700px;border:2px black solid" id="carfax"></iframe>
<input id="carfaxbutton" onclick="closeCarfax()" style="z-index:999;display:none;height:39px; font-size:large;color:red;font-weight:bold;position:absolute;top:0px;left:0px;" type="button" value="Close Carfax History">
<input id="lastguide" type="hidden">
<input type="hidden" id="requiremileagetoprintro" value="<%=requiremileagetoprintro%>">
<input name="goback" id="goback" type="hidden" value="no" >
<input name="gobackurl" id="gobackurl" type="hidden" >
<input id="backbutton" style="font-size:18px;z-index:999;height:5%;color:red;font-weight:bold;display:none;position:absolute;top:5px;left:5%;" onclick="showPOs('hide')" type="button" onclick="" value="Back To PO Managment">
<iframe id="pos" style="width:90%;position:absolute;left:5%;top:5%;height:90%;background-color:white;display:none;z-index:1000"></iframe>
<div style="display:none" id="spouseinfo">
<br/><br/>
	Spouse Name: <%=rs("spousename")%><br><br><br>
	Spouse Work Phone: <%=rs("spousework")%><br><br><br>
	Spouse Cell Phone: <%=rs("spousecell")%><br><br><br>
<input onclick="hideSpouse()" name="Button1" type="button" value="Close" >
</div>
<input type="hidden" id="changes" name="changes" value="no">
<input type="hidden" id="inclprint" name="inclprint" value="">
<input type="hidden" id="timeclocklaborid" name="timeclocklaborid" value="">
<input type="hidden" id="timerunning" name="timerunning" value="no">
<div style="position:absolute;top:100px;width:400px;height:314px; left:40%;background-color:white;padding:25px;display:none;z-index:1005" id="sendemail">
	Please verify customer email. You can change it if necessary<br>
<input type="text" id="curremailaddress" value="<%=cemail%>" style="width: 261px" ><br><br>
Enter a message to the customer:<br>
<textarea style="width:300px;height:100px;" id="emailinvoicemessage">Your repair invoice from <%=request.cookies("shopname")%> is attached.  Thank you for your business!</textarea>

	<br><input id="updateemail" value="on" name="Checkbox1" type="checkbox"> 
	Update customer info with this email address<br>
	<br>&nbsp;<input onclick="redirectWithEmail()" name="Button4" type="button" value="Send Email"> <input onclick="document.getElementById('hider').style.display='none';document.getElementById('sendemail').style.display='none'" name="Button1" type="button" value="Cancel" ></div>
<%if esign = "yes" then%>
<style type="text/css">
#printdrop{display:none;position:absolute;width:100%;background-color:#F0F0F0;z-index:1001;font-weight:700;left:0;top:50px;padding:0;border:3px #06C solid}.printmenu{padding:5px;cursor:pointer;text-align:center;width:20%;border:3px #06c solid}.printmenu:hover{color:#000;background-color:#fff}.auto-style1{border-style:solid}</style>

<table style="width:60%;margin-left:20%" id="printdrop" cellpadding="2" cellspacing="0" class="auto-style1">
	<tr>
		<td onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none';document.body.onunload='';saveAll('ro.asp?printro=yes&','','','','','')" class="printmenu"><img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br>Print RO&nbsp;</td>
		<td onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none';document.body.onunload='';saveAll('ro.asp?printwo=yes&','','','','','')" class="printmenu"><img src="newimages/wrench.png" height="30" width="27"><br/>Print Tech Work Order&nbsp;</td>
		<td onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none'" class="printmenu"><img src="newimages/close-icon.png" height="30" width="27"><br/>Close&nbsp;</td>
	</tr>
	<tr>
		<td colspan="5">
		
		<div style="width:100%;" id="printedrolist">
		Any E-Signed RO's Are listed here:<br><br>
		<%
		'on error resume next
		set fso = server.createobject("scripting.filesystemobject")
		folpath = "d:\savedinvoices\" & shopid & "\" & request("roid")
		newpath = server.mappath(folpath)
		'response.write newpath
		if fso.folderexists(newpath) then
			
			'response.write newpath
			fol = fso.getfolder(newpath)
			if fso.folderexists(fol) then
				set mfol = fso.getfolder(fol)
				c = 1
				for each f in mfol.files
					mmod = c mod 2
					if mmod = 2 then
						bgcolor = "#ffc"
					else
						bgcolor = "#fff"
					end if
					filename = ucase(f.name)
					far = split(filename,"_")
					' 2-7
					filedate = far(2) & "/" & far(3) & "/" & far(4) & " at " & far(5) & ":" & far(6) & ":" & far(7)
					'response.write "filename:" & filename
					response.write "<div onclick='reprintRO(" & chr(34) & f.name & chr(34) & ")' style='text-align:center;padding:4px;color:#3366CC;cursor:pointer;background-color:" & bgcolor & "'>Invoice printed and dated " & filedate & "</div><br>"
				next
			else
				response.write "No previously signed RO's on file"
			end if
		end if
		%>
		</div>
		</td>
	</tr>
</table>
<%else%>
<style type="text/css">
#printdrop{list-style-type:none;display:none;position:absolute;width:150px;background-color:#F0F0F0;z-index:1001;font-weight:700;left:45%;top:100px;padding:5px;border:thin gray solid}li{padding:5px;cursor:pointer;text-align:center;border-bottom:3px #06C solid}li:hover{color:#000;background-color:#fff}
</style>

<ul id="printdrop" style="">
<li onclick="printROwithSave()"  style=""><img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br/>
Print RO</li>
<li onclick="showEmailAddress('no')" style="">

<img src="newimages/emailicon.gif" height="30" width="27"><br/>Email RO</li>
<li onclick="showEmailAddress('yes')" style="">

<img alt="" height="30" src="newimages/printer_icon.gif" width="33">+<img src="newimages/emailicon.gif" height="30" width="27"><br/>
Print &amp; Email RO</li>
<li  onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none';document.body.onunload='';saveAll('ro.asp?printwo=yes&','','','','','')" style="">

<img src="newimages/wrench.png" height="30" width="27"><br/>Print Tech Work Order</li>

<li onmouseout="this.style.backgroundColor='#F0F0F0';this.style.color='black'" onmouseover="this.style.backgroundColor='white';this.style.color='black'" onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none'" style="">

<img src="newimages/close-icon.png" height="30" width="27"><br/>Cancel</li>

</ul>
<%end if%>

<input id="dfield" name="dfield" type="hidden" ><input name="caldatefield" type="hidden" id="caldatefield">
<div style="z-index:1000;display:none;width:60%;position:absolute;left:20%;top:100px;height:500px;background-color:white;padding:20px;" id="reminders">
	<strong>Service Reminders - <span onclick="addReminder();document.remform.reminder.focus()" style="color:#0258BD;cursor:pointer;padding:10px;border:medium black outset;background-color:#3366CC;color:white;border:thin navy outset">
	Add Reminder</span>
	<input onclick="closeReminder()" style="font-size:xx-small;width:40px;position:absolute;top:15px;right:15px;cursor:pointer;height:30px" name="Button1" type="button" value="Close" >
	</strong><br><br>
	<table style="width: 100%" cellpadding="4" cellspacing="0" >
		<tr>
			<td class="style10" style="height: 46; width: 75%;"><strong>
			Reminder&nbsp;</strong></td>
			<td class="style10" style="height: 46; width: 30%;"><strong>Date&nbsp;</strong></td>
			<td class="style10" style="height: 46; width: 30%;"><strong>Delete&nbsp;</strong></td>
		</tr>
		<%
		stmt = "select * from reminders where shopid = '" & shopid & "' and customerid = " & rs("customerid") & " and vehid = " & rs("vehid") _
		& " order by reminderdate asc"
		'response.write stmt
		set remrs = con.execute(stmt)
		if not remrs.eof then
			do until remrs.eof
		%>
		<tr>
			<td style="width: 75%; height: 24px"><%=pcase(remrs("reminderservice"))%>
			&nbsp;</td>
			<td style="height: 24px"><%=remrs("reminderdate")%>&nbsp;</td>
			<td onclick="deleteReminder('<%=remrs("id")%>')" style="height: 24px;color:#3366CC;cursor:pointer">
			Delete&nbsp;</td>
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

<div style="z-index:1000;display:none;width:80%;position:absolute;left:10%;top:100px;height:400px;background-color:white;padding:20px;border:medium black outset;overflow-y:scroll" id="addreminder">
	<form action="postReminder()" method="post" name="remform">
	<input name="customerid" type="hidden" value="<%=rs("customerid")%>" ><input name="vehid" type="hidden" value="<%=rs("vehid")%>">
	<input name="roid" type="hidden" value="<%=request("roid")%>">
	<strong>Enter the Service Reminder and Date<br><span style="font-size:small">
		(Reminders will be sent via email if the customer has an email address 
		on file)</span></strong>&nbsp;<table style="width: 100%" cellpadding="4" cellspacing="0">
		<tr>
			<td class="style10" style="height: 44; width: 50%;">Reminder 
			Description</td>
			<td class="style10" style="height: 44; width: 20%;">Date Due&nbsp;</td>
			<td class="style34" style="height: 44; width: 10%;">Recurring?</td>
			<td class="style34" style="height: 44; width: 20%;">Interval Days<br>
			<span class="style35">How Often to Send Reminder</span></td>
		</tr>
		<tr>
			<td style="width: 278px">
			<input name="reminder" type="text" style="width: 404px; text-transform:capitalize" >&nbsp;</td>
			<td>&nbsp;<input onfocus="show_cal(this,'demo1');" id="demo1" name="reminderdate" style="width: 120px" type="text"> <img src="javascripts/images2/cal.gif" onclick="show_cal(this,'demo1');" style="cursor:pointer"></td>
			<td class="style36"><select name="recurring" style="width: 69px">
				<option value="Select">Select</option>
				<option value="yes">Yes</option>
				<option value="no">No</option>
			</select>&nbsp;</td>
			<td class="style36">
			<input name="interval" type="text" style="width: 61px" >&nbsp;</td>
		</tr>
	</table>
	<input onclick="saveReminder()" name="Button1" type="button" value="Add Reminder" > <input onclick="closeAddReminder()" name="Button1" type="button" value="Cancel" >
	</form>
</div>

<table id="toptable" border=2 cellpadding=0 cellspacing=0 width=100% bordercolor="#0066CC" style="height:100%"><tr>
	<td valign="top" style="">
<form name="theform" method="post" action="1238savero.asp">
<input type="hidden" id="changeorder" name="changeorder" value="">
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

<div style="display:none" id="recrepairs">

	<table cellpadding="2" style="width:100%">
		<tr>
			<td style="color:white;font-weight:bold;width:100%" colspan="4" class="style18">
<input id="b1" style="cursor:pointer" name="Button16" type="button" onclick="printDiv('recrepairs')" value="Print" ><input id="b2" style="cursor:pointer" name="Button17" type="button" onclick="closeRecRepairs()" value="Close" ><input id="b3" style="cursor:pointer" name="Button18" type="button" onclick="addRecRepair()" value="Add Recommended Repair" ></td>
		</tr>
		<tr>
			<td style="font-weight:bold" colspan="3">
	Recommended Repairs:</td>
		</tr>
		<tr>
			<td style="background-color:maroon;color:white"></td>
			<td style="background-color:maroon;color:white;font-weight:bold;text-align:left">
			RECOMMENDATIONS&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold;text-align:right">
			Repair Cost&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold;text-align:right">
			Restore/Delete&nbsp;</td>
		</tr>
		<%
		stmt = "select * from recommend where shopid = '" & shopid & "' and roid = " & request("roid") 
		'response.write stmt
		set recrs = con.execute(stmt)
		if not recrs.eof then
			do until recrs.eof
				
		%>
		<tr>
			<td></td>
			<td id="desc" style="font-weight:bold;text-align:left"><%=recrs("desc")%></td>
			<td id="amount" style="font-weight:bold;text-align:right"><%=formatnumber(recrs("totalrec"),2)%></td>
			<td style="font-weight:bold;text-align:right">
			<input onclick="editRec('<%=recrs("id")%>','<%=recrs("desc")%>','<%=formatnumber(recrs("totalrec"),2)%>')" name="Button1" type="button" value="Edit">
			<input onclick="document.getElementById('recrepairs').style.display='none';document.body.onunload='';document.theform.action='savero.asp?recid=<%=recrs("id")%>&comid=<%=recrs("comid")%>&comstat=Approved';saveAll('ro.asp?a=n&','','','','','')" name="Button1" type="button" value="Restore" > 
			<input onclick="deleteRec('<%=recrs("id")%>')" name="Button1" type="button" value="Delete" >
			</td>
		</tr>
		<%
				recrs.movenext
			loop
		else
		%>
		<tr>
			<td style="font-weight:bold;text-align:left" colspan="4">No Recommendations</td>
		</tr>
		<%
		end if
		
		' now check for recommended from previous RO's
		' first get a list of all roid's for this vehicle
		
		stmt = "select roid from repairorders where roid <> " & request("roid") & " and shopid = '" & request.cookies("shopid") & "' and vehid = " & rs("vehid") & " order by datein desc"
		'response.write stmt & "<br>"
		set roidrs = con.execute(stmt)
		if not roidrs.eof then
			do until roidrs.eof
				roidlist = roidlist & roidrs("roid") & ","
				roidrs.movenext
			loop
		end if
		'response.write "<BR>" & roidlist
		if len(roidlist) > 3 and instr(roidlist,",") then
			stmtadd = stmtadd & "(roid = 1"
			rolistar = split(roidlist,",")
			for j = lbound(rolistar) to ubound(rolistar)-1
				stmtadd = stmtadd & " or roid = " & rolistar(j)
			next
			stmt = "select * from recommend where shopid = '" & shopid & "' and " & stmtadd & ")" 
			'response.write  "<BR>" & stmt
			set recrs = con.execute(stmt)
			if not recrs.eof then
		%>
		<tr>
			<td style="text-align:left;color:white;background-color:maroon" colspan="4">*************** DECLINED REPAIRS FROM PREVIOUS REPAIR ORDERS **************</td>
		</tr>
		<tr>
			<td style="text-align:left;color:black;background-color:white" colspan="4"><input onclick="restoreAllRecs()"  name="Button1" type="button" value="Add all Checked Recommendations" ></td>
		</tr>
		
		<%
				do until recrs.eof
					stmt = "select customer,vehinfo from repairorders where shopid = '" & recrs("shopid") & "' and roid = " & recrs("roid")
					set rec2rs = con.execute(stmt)
					if not rec2rs.eof then
						rec2customer = rec2rs("customer")
						rec2veh = rec2rs("vehinfo")
					end if
					set rec2rs = nothing
		%>
		<tr>
			<td><input name="rrcb<%=recrs("id")%>" type="checkbox" id="rrcb<%=recrs("id")%>" ></td>
			<td id="desc" style="font-weight:bold;text-align:left"><%="RO#" & recrs("roid") & "-" & ucase(recrs("desc"))%></td>
			<td id="amount" style="font-weight:bold;text-align:right"><%=formatnumber(recrs("totalrec"),2)%></td>
			<td style="font-weight:bold;text-align:right">
			<input onclick="addToRO('<%=recrs("id")%>','<%=recrs("comid")%>')" name="Button1" type="button" value="Add to RO" > 
			<input onclick="deleteRec('<%=recrs("id")%>')" name="Button1" type="button" value="Delete" >
			</td>
		</tr>
		<%
					recrs.movenext
				loop
			end if
		end if
		%>

	</table>
<div style="text-align:left;height:400px">
</div>

</div>
<div style="display:none" id="customvehiclefields">
<%

' write out the fields for tire pressures and tread depth
%>

	<table style="width: 100%" cellpadding="3" cellspacing="0">
		<tr>
			<td class="style36" colspan="5"><strong>Tire Pressure and Tread Depth</strong></td>
		</tr>
		<tr>
			<td class="style37"><strong>Tire Pressure IN</strong></td>
			<td class="style37"><strong>LF:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureinlf")%>" id="tirepressureinlf" name="tirepressureinlf" style="width: 36px" type="text">PSI</strong></td>
			<td class="style37"><strong>RF:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureinrf")%>" id="tirepressureinrf" name="tirepressureinrf" style="width: 36px" type="text">PSI</strong></td>
			<td class="style37"><strong>LR:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureinlr")%>" id="tirepressureinlr" name="tirepressureinlr" style="width: 36px" type="text">PSI</strong></td>
			<td class="style37"><strong>RR:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureinrr")%>" id="tirepressureinrr" name="tirepressureinrr" style="width: 36px" type="text">PSI</strong></td>
		</tr>
		<tr>
			<td class="style38"><strong>Tire Pressure OUT</strong></td>
			<td class="style38"><strong>LF:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureoutlf")%>" id="tirepressureoutlf" name="tirepressureoutlf" style="width: 36px" type="text">PSI</strong></td>
			<td class="style38"><strong>RF:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureoutrf")%>" id="tirepressureoutrf" name="tirepressureoutrf" style="width: 36px" type="text">PSI</strong></td>
			<td class="style38"><strong>LR:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureoutlr")%>" id="tirepressureoutlr" name="tirepressureoutlr" style="width: 36px" type="text">PSI</strong></td>
			<td class="style38"><strong>RR:<input onkeyup="chgDB(this.name)" value="<%=rs("tirepressureoutrr")%>" id="tirepressureoutrr" name="tirepressureoutrr" style="width: 36px" type="text">PSI</strong></td>
		</tr>
		<tr>
			<td class="style39"><strong>Tread Depth</strong></td>
			<td class="style39"><strong>LF:<input onkeyup="chgDB(this.name)" value="<%=rs("treaddepthlf")%>" id="treaddepthlf" name="treaddepthlf" style="width: 36px" type="text">/32"</strong></td>
			<td class="style39"><strong>RF:<input onkeyup="chgDB(this.name)" value="<%=rs("treaddepthrf")%>" id="treaddepthrf" name="treaddepthrf" style="width: 36px" type="text">/32"</strong></td>
			<td class="style39"><strong>LR:<input onkeyup="chgDB(this.name)" value="<%=rs("treaddepthlr")%>" id="treaddepthlr" name="treaddepthlr" style="width: 36px" type="text">/32"</strong></td>
			<td class="style39"><strong>RR:<input onkeyup="chgDB(this.name)" value="<%=rs("treaddepthrr")%>" id="treaddepthrr" name="treaddepthrr" style="width: 36px" type="text">/32"</strong></td>
		</tr>
	</table>

<%

if len(rs("customvehicle1label")) > 0 then
response.write rs("customvehicle1label") & ": " & rs("customvehicle1") & "<br>"
end if
if len(rs("customvehicle2label")) > 0 then
response.write rs("customvehicle2label") & ": " & rs("customvehicle2") & "<br>"
end if
if len(rs("customvehicle3label")) > 0 then
response.write rs("customvehicle3label") & ": " & rs("customvehicle3") & "<br>"
end if
if len(rs("customvehicle4label")) > 0 then
response.write rs("customvehicle4label") & ": " & rs("customvehicle4") & "<br>"
end if
if len(rs("customvehicle5label")) > 0 then
response.write rs("customvehicle5label") & ": " & rs("customvehicle5") & "<br>"
end if
if len(rs("customvehicle6label")) > 0 then
response.write rs("customvehicle6label") & ": " & rs("customvehicle6") & "<br>"
end if
if len(rs("customvehicle7label")) > 0 then
response.write rs("customvehicle7label") & ": " & rs("customvehicle7") & "<br>"
end if
if len(rs("customvehicle8label")) > 0 then
response.write rs("customvehicle8label") & ": " & rs("customvehicle8") & "<br><br>"
end if

%>
<input name="Button1" type="button" onclick="hideVehicleFields()" value="Close" />
</div>
<input type="hidden" name="x" value="" >
<input type="hidden" name="oldstatus" value="<%=rs("status")%>" >
 <input type="hidden" name="dir" value><input type="hidden" name="itype" value><input type="hidden" name="pid" value><input type="hidden" name="sid" value><input type="hidden" name="lid" value>
  
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
   <tr>
    <td valign="top">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
       <tr>
        <td valign="top" width="100%">
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
           <tr>
            <td valign="top" height="22" >
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
               <td width="100%" align="center" class="style10" ><b><i>
               <span style="float: left;">
               <img onclick="showCFHistory('<%=rs("vin")%>','<%=rs("vehlicense")%>','<%=rs("vehstate")%>')"  alt="" src="carfaxshclogo-28.jpg" height="28" style="cursor:pointer" width="89">
               <img onclick="openTPMS()" alt="" src="newimages/TPMS SPEED BUTTON LOGO.jpg" style="cursor:pointer">
               </span>
               <span style="display:none;" id="gp"></span>
               <font color="#FFFF00" size="5">RO#&nbsp;<%=roid%></font></i></b><%if useim = "yes" then%><button type="button" id="chat" onclick="showHide()" style="position:absolute;right:160px;top:0px;cursor:pointer;border:1px black solid;height:37px; font-weight:bold;background-color:yellow;color:black;border-radius:4px;font-size:small; width: 87px;" >Chat</button><%end if%><img style="position:absolute;right:5px;top:0px;" alt="" height="37" src="newimages/shopbosslogo-ro.png" width="157"></td>
              </tr>
              <tr>
               <td width="100%" align="left" bgcolor="#0066CC">

             
               <table style="width: 100%">
				   <tr>
					   <td onclick="document.theform.Status.focus();document.body.onunload='';saveAll('ro.asp?a=n&','','','','','')"  class="menustyle">
					   <img align="middle" alt="" height="30" src="newimages/icon-calculator.jpg" width="31"><br>
					   Calculate</td>
					   <td onclick="findPosPrint(this)" valign="middle" class="menustyle">
					   <img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br>
					   Print RO</td>
					   <td onclick="showInspection()" class="menustyle">
					   <img alt="" height="29" src="newimages/inspection.gif" width="40"><br>
					   Inspection</td>
					   <td onclick="showHistory()" class="menustyle">
					   <img alt="" height="30" src="newimages/historyx30.png" width="30"><br>
					   History</td>
					   <td onclick="uploadFiles()" class="menustyle">
					   <img alt="" height="30" src="newimages/upload.png" width="39"><br>
					   Upload/View Pics</td>
					   <td style="z-index:999" onclick="showPOs('')" class="menustyle">
					   <img alt="" height="30" src="newimages/icon_salesorders.gif" width="31"><br>
					   					   PO Number</td>
					   <td style="z-index:999;" onclick="postAllParts()" class="menustyle">
					   <img alt="" height="30" src="newimages/accticonlarge.png" width="30"><br>
					   					   Post All Parts to Acct</td>
					   <td onclick="setReminder()" class="menustyle">
					   <img alt="" height="30" src="newimages/reminder.png" width="30"><br>
					   Set/View Reminders</td>
					   <td onclick="<%if lcase(request.cookies("sendupdates")) = "yes" then response.write "showUpdateWindow('y')"%>" class="menustyle">
					   <img alt="" height="30" src="newimages/alert.png" width="30"><br>
					   Send Update</td>
					   
					   <td onclick="document.body.onunload='';saveAll('complaintaddupdate.asp?a=b&','','','','','')" class="menustyle">
					   <img alt="" height="30" src="newimages/vintage-car-icon-ro.png" width="45">&nbsp;<br>
					   &nbsp;Vehicle Issues</td>
					   <td onclick="showRecRepairs()" class="menustyle">
					   <img alt="" height="30" src="newimages/wrench.png" width="45">&nbsp;<br>
					   &nbsp;Recommended Repairs</td>
					   <td onclick="showOrderSelect()" style="background-color:#CC0000;color:white" class="menustyle">

					   <img alt="" height="30" src="newimages/calculator-icon-2.png" width="30"><br>
					   Estimating &amp; Parts Ordering</td>

					   <td onclick="document.theform.Status.focus();saveAll('wip.asp?a=n&','','','','','')" class="menustyle">
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
         
         <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
           <td width="34%" valign="top" bgcolor="#ffffff" height="170">
           <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" height="100%" class="style16">
             <tr>
              <td ondblclick="findPosPrint(this)" id="temp" width="10%" align="right" valign="bottom" class="style14">
              RO/Tag #</td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15">
			  <span id='tagnumber'><%=ucase(rs("tagnumber"))%></span>&nbsp;&nbsp;
			  <%
			  if len(rs("tagnumber")) > 0 then
			  %>
			  <span id="deltagnumber" onclick="document.theform.dbtagnumber.value='';document.getElementById('tagnumber').style.display='none';document.getElementById('deltagnumber').style.display='none'" style="font-weight:bold;font-size:x-small;color:blue;cursor:pointer">
			  (Delete Tag #)</span></td>
              <%
              end if
              %>
             </tr>
             <tr>
              <td ondblclick="" width="10%" align="right" valign="bottom" class="style14" style="height: 21px">
              <font size="2" face="Arial">Customer:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15" style="height: 21px"><font size="2" face="arial"><%=server.htmlencode(rs("Customer"))%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td ondblclick="document.body.onunload='';saveAll('ro.asp?printro=yes&','','','','','')" width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Address:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=rs("CustomerAddress")%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">City,State,Zip:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=rs("CustomerCSZ")%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Home:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=homephone%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Work:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=workphone%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              <font size="2" face="Arial">Cell:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=cellphone%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              Email: </td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><%=cemail%>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
              Bus. Contact: </td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><%=rs("contact")%>
			  &nbsp;</td>
             </tr>
            </table></td>
           <td width="33%" valign="top" bgcolor="#FFFFFF" height="170">
           <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" height="100%" class="style16">
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Vehicle:</font></td>
              <td style="overflow:hidden" width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial">
              <span style="color:#3366CC;text-decoration:underline;cursor:pointer" onclick="showCFHistory('<%=rs("vin")%>','<%=rs("vehlicense")%>','<%=rs("vehstate")%>')">
              <%=left(rs("VehInfo"),27)%>
              </span>
              </font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14" style="height: 23px">
              <font size="2" face="Arial">License:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 23px"><font size="2" face="arial"><%=rs("VehLicNum")%>
			  &nbsp;&nbsp;&nbsp;Fleet #<%=rs("fleetno")%></font></td>
             </tr>
             <tr>
              <td onclick="saveAll('pdfinvoices/1073/printpdfronew.asp?roid=<%=request("roid")%>&','','','','');"" width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">VIN:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("Vin")%></font>
			  &nbsp;</td>
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
              <td width="10%" align="right" valign="top" class="style14" style="height: 22px">
              <font size="2" face="Arial">Engine/Drive:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 22px"><font size="2" face="arial"><%=rs("VehEngine")%>
			  &nbsp;-&nbsp;<%=rs("Cyl")%>Cyl&nbsp;-&nbsp;<%=rs("VehTrans")%>&nbsp;-&nbsp;<%=rs("DriveType")%></font></td>
             </tr>
             <tr>
              <td colspan="2" style="cursor:pointer;background-color:white" width="20%" valign="top" class="style24">
			  <input onclick="showSpouse()" name="button15" type="button" value="Additional Customer Info" style="width: 174px">
			  <input name="Button1" onclick="showVehicleFields()" style="width: 174px" type="button" value="Additional Vehicle Fields">
			  <br><br>
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
			  	  In:<input onfocus="show_cal(this,'DateIn');document.getElementById('caldatefield').value='DateIn'" id="DateIn" size="10" onblur="" name="DateIn" value="<%=rs("datein")%>" type="text" > 
				  Status:<input onchange="document.theform.dbFinalDate.value=this.value;alert(document.theform.dbFinalDate.value)" id="StatusDate" onfocus="show_cal(this,'StatusDate');document.getElementById('caldatefield').value='StatusDate';document.theform.overridedate.value='yes'" size="10" onblur="document.theform.dbFinalDate.value=this.value;alert(document.theform.dbFinalDate.value)" name="StatusDate" value="<%=rs("statusdate")%>" type="text" >
			  </div>
			  </td>
             </tr>
             <tr>
             
              <td ondblclick="alert(document.theform.dbStatusDate.value+':'+document.theform.dbDateIn.value)" width="10%" align="right" valign="top" height="44" style="height: 22px" class="style14">
			  <input type="hidden" name="overridedate" value="no">Status</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="44" style="height: 22px" class="style15">
			  <select style="height:18px;" onchange="setStatusDate();chgDB(this.name);setTimeout('closeRO()',500)" size="1" name="Status" >
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
              <td width="10%" align="right" valign="top" class="style14" style="height: 31px"><font size="2" face="arial">
			  Comments:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 31px"><font size="2" face="arial">
              <textarea onfocus="javascript:this.style.backgroundColor='#FFFFCC'" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white';" name="Comments" cols="37" style="padding: 1 4; text-transform: uppercase; font-size: 8pt; font-family: Arial; height: 39px;" tabindex="21"><%=rs("Comments")%></textarea></font></td>
             </tr>
             <tr>
              <td ondblclick="saveAll('pdfinvoices/printpdfro-new.asp?roid=<%=request("roid")%>&','','','','')" width="10%" align="right" valign="top" class="style14" style="height: 22px"><font size="2" face="arial">
			  Writer:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 22px">
               <select onchange="javascript:chgDB(this.name);" id="Writer" size="1" name="Writer">
                <option selected value="<%=rs("Writer")%>"><%=rs("Writer")%></option>
                <%
				estmt = "Select EmployeeFirst, EmployeeLast from employees where shopid = '" & shopid & "' and Active = 'Yes'"
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
              <td ondblclick="saveAll('pdfinvoices/printpdfro-new.asp?roid=<%=request("roid")%>&','','','','')" width="10%" align="right" valign="top" class="style14" style="height: 22px">
			  PO Number</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" class="style15" style="height: 22px">
               <font size="2" face="arial"><b>
               <input onkeyup="chgDB(this.name);" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" id="MilesOut0" name="ponumber" value='<%=rs("ponumber")%>' style="border:1px solid #BBD7F2; height:18px; font-family:Arial; font-size:10pt; font-variant:small-caps; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1" size="9"></b></font></td>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14"><font size="2" face="arial">
			  Source:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15"><font size="2" face="arial">
               </font><select size="1" onchange="javascript:chgDB(this.name);" id="Source" name="Source" >

                <%
				srcstmt = "select Source from source where shopid = '" & shopid & "'"
				set srcrs = con.execute(srcstmt)
				if not srcrs.eof then
					while not srcrs.eof
						if srcrs("source") = rs("source") then
							s = " selected='selected' "
						else
							s = ""
						end if
						response.write "<option " & s & " value='" & srcrs("Source") & "'>" & srcrs("Source") & "</option>"
						srcrs.movenext
					wend
				else
					response.write "<option value='None Entered'>None Entered</option>"
				end if
                %>
               </select></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14">
               <p align="right" class="style13"><font size="2" face="arial">
			   Type:</font></td>


              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15"><select size="1" onchange="javascript:chgDB(this.name);" onblur="javascript:chgDB(this.name);this.style.backgroundColor='white'" id="ROType0" name="ROType" size="20">
                <option selected value="<%=rs("ROType")%>"><%=rs("ROType")%></option>
                <%
				srcstmt = "select ROType from rotype where shopid = '" & shopid & "'"
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
		   <div id="issues" >
           <!-- #include file=1238complaintlayout.asp -->
		   </div>

           </td>
           <td width="33%" valign="top">
           <table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse; height:30px;" bordercolor="#111111">
             <tr>
              <td style="cursor:pointer" id="revcell" Width="25%" align="center" onclick="revChg()" class="style9">
			  <img alt="" height="30" src="newimages/desktop.png" width="30"><br>
			  Revisions</td>
              <td style="cursor:pointer" id="feecell" width="25%" align="center" onclick="feesChg()" class="style9">
			  <img alt="" height="30" src="newimages/discount_icon.png" width="28"><br>
			  Fees/Taxes</td>
              <td style="cursor:pointer" id="warrcell" width="25%" align="center" onclick="warrChg()" class="style9">
			  <img alt="" height="30" src="newimages/warricon2.png" width="32"><br>
			  Warranty</td>
              <td style="cursor:pointer" id="pmtscell" width="25%" align="center"onclick="pmtsChg()" class="style9">
			  <img alt="" height="30" src="newimages/payment.png" width="30"><br>
			  Payments</td>
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
                 <td width="50%" colspan="2" class="style30">
                  Promise Date:
                 </td>
                 <td width="50%" colspan="2" bgcolor="#FFFFFF" class="style15">
				 <input style="height: 18; font-family: Arial; font-variant: small-caps; font-size: 8pt; width: 105px;" onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="datetimepromised" size="9" value='<%=rs("datetimepromised")%>'></td>
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
                 <td width="25%" align="right" class="style17"><font size="2">
				 Amt:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height:18; font-variant: small-caps; font-family: Arial; font-size: 8pt; text-align: Right; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1Amt" size="9" value="<%=formatnumber(rs("Rev1Amt"),2)%>" tabindex="10"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Amt:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Amt" size="9" value="<%=formatnumber(rs("Rev2Amt"),2)%>" style="font-family: Arial; font-variant: small-caps; font-size: 8pt; text-align: Right; height:18; " tabindex="15"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onkeyup="chgDB(this.name);" onblur="chgDB(this.name)" type="text" name="Rev1Date" size="9" value="<%=rs("Rev1Date")%>" tabindex="11"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Date" size="9" value="<%=rs("Rev2Date")%>" tabindex="16"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onkeyup="chgDB(this.name);" onblur="chgDB(this.name)" type="text" name="Rev1Phone" size="9" value="<%=rs("Rev1Phone")%>" tabindex="12"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Phone" size="9" value="<%=rs("Rev2Phone")%>" tabindex="17"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1Time" size="9" value="<%=rs("Rev1Time")%>" tabindex="13"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2Time" size="9" value="<%=rs("Rev2Time")%>" tabindex="18"></font></td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-variant: small-caps; font-family: Arial; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onkeyup="chgDB(this.name);" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev1By" size="9" value="<%=rs("Rev1By")%>" tabindex="14"></font></td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><font size="2">
				 <input style="height: 18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" name="Rev2By" size="9" value="<%=rs("Rev2By")%>" tabindex="19"></font></td>
                </tr>
               </table></td>
             </tr>
             <tr>
              <td id="fees" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td ondblclick="calcPercents()" align="right" class="style14" style="height: 24px; width: 41%;"><font size="2">
				 Haz. Waste Fee:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" style="height: 24px" colspan="2">
                  $ <input style="height:18; font-family: Arial; font-variant: small-caps; text-align:right; font-size: 8pt; width:60px; " onblur="setZero(this.name,this.value);chgDB(this.name);ttlFees();this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="HazardousWaste" size="9" value="<%=formatdollar(rs("HazardousWaste"))%>">
                 </td>
                </tr>
                 <%
                 'get userfees from company
                 
                 if len(rs("userfee1label")) > 0 then
               		userfee1desc = rs("userfee1label") & ":"
                 	userfee1amount = formatnumber(rs("userfee1"),2)
                 	userfee1percent = formatnumber(rs("userfee1percent"),3)
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
                 <td align="right" class="style14" style="width: 41%"><%=userfee1desc%>
				 &nbsp;</td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2"><%=userfee1display%></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%"><%=userfee2desc%>
				 &nbsp;</td>
                   <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2"><%=userfee2display%></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%"><%=userfee3desc%>
				 &nbsp;</td>

                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2"><%=userfee3display%></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%"><font size="2">
				 Storage Fee:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2">
				 $
				 <input style="height:18px; font-family: Arial; font-variant: small-caps; font-size: 8pt;text-align: right; width:60px;" onblur="setZero(this.name,this.value);chgDB(this.name);ttlFees();this.style.backgroundColor='white'" type="text" name="storagefee" size="9" value='<%=formatdollar(rs("storagefee"))%>'></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%"><font size="2">
				 Parts Tax Rate:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2"><input style="height:18px; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="TaxRate" size="6" value="<%=rs("TaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%">Labor Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2">
                 <input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="LaborTaxRate" size="6" value="<%=rs("LaborTaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%">Sublet Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" colspan="2">
                 <input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" onkeyup="chgDB(this.name);" type="text" name="SubletTaxRate" size="6" value="<%=rs("SubletTaxRate")%>"><span class="style22">%</span></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%"><font size="2">
				 Discount Amt:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" style="width: 0%">
				 &nbsp;<input name="discpercent" onblur="" style="width: 23px" type="text" value="0">%<input onclick="discountPercent(document.theform.discpercent.value)" style="font-size:xx-small" name="Button3" type="button" value="Calc"></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15" style="width: 25%">
				 $<input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; width:60px; " onblur="setZero(this.name,this.value);chgDB(this.name);this.style.backgroundColor='white'" type="text" onkeyup="chgDB(this.name);" style="text-align:right" name="DiscountAmt" size="9" value="<%=formatdollar(rs("DiscountAmt"))%>"></td>
                </tr>
                <tr>
                 <td align="right" class="style14" style="width: 41%">Discount 
				 Taxable:&nbsp; </td>
                 <td align="right" bgcolor="#FFFFFF" class="style15" style="width: 100%" colspan="2">
                 <%
                 if lcase(rs("discounttaxable")) = "yes" then
                 	yselect = "selected='selected'"
                 	nselect = ""
                 else
                 	yselect = ""
                 	nselect = "selected='selected'"
                 end if
                  ' ***********************   DISCOUNT TAXABLE  *******************************
                 %>
                 <select onchange="chgDB(this.name)" name="discounttaxable" style="width: 65px">
					<option <%=yselect%> value="yes">Yes</option>
					<option <%=nselect%> value="no">No</option>
				</select></td>
                </tr>
                </table></td>
             </tr>
             <tr>
              <td id="warr" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td  width="50%" align="right" class="style17"><font size="2">
				 Warranty Months:&nbsp;</font></td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="height:18; font-family: Arial; font-variant: small-caps; font-size: 8pt; " onfocus="javascript:this.select();this.style.backgroundColor='#FFFFCC'" onblur="chgDB(this.name);this.style.backgroundColor='white'" type="text" onkeyup="chgDB(this.name);" style="text-align:right" name="WarrMos" size="7" value="<%=rs("WarrMos")%>"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17"><font size="2">
				 Warranty Miles:&nbsp;</font></td>
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
               <!-- alternate table here -->
              <table width="100%" border="1" cellspacing="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
                 <td class="style27" style="width: 33%">
				 Paid with</td>
                 <td width="33%" class="style27" style="width: 33%">
				 Date</td>
                 <td width="33%" align="right" class="style14" style="width: 33%">
				 Amount</td>
				 </tr>
				 <%
				 stmt = "select * from accountpayments where shopid = '" & shopid _
				 & "' and roid = " & request("roid")
				 set payrs = con.execute(stmt)
				 pcntr = 1
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
				 		pcntr = pcntr + 1
				 		payrs.movenext
				 	loop
				 	if pcntr < 3 then
				 		for j = 1 to 3-pcntr
				 	%>
	                 <tr>
	                 <td class="style29" style="width: 33%">
					 &nbsp;</td>
	                 <td width="33%" class="style29" style="width: 33%">
					 &nbsp;</td>
	                 <td width="33%" align="right" class="style15" style="width: 33%">
					 &nbsp;</td>
					 </tr>

				 	<%
				 		next
				 	end if
				 else
				 %>
                 <tr>
                 <td colspan="3" class="style29" style="width: 33%">No Payments 
				 posted<br><br><br>
				 &nbsp;</td>
				 </tr>
				 <%
				 end if
				 if merchantaccount = "no" then
	                 paystmt = "select sum(`amt`) as payments from accountpayments where roid = " & request("roid") & " and shopid = '" & shopid & "'"
	                 set payrs = con.execute(paystmt)
	                 if isnull(payrs("payments")) then
	                 	otherpayments = 0
	                 else
	                 	otherpayments = payrs("payments")
	                 end if
	                 'response.write formatdollar(otherpayments)
	                 set payrs = nothing
	                 'response.write otherpayments & ":" & rs("amtpaid1") & ":" & rs("amtpaid2")
	                 customerpayments = otherpayments '+ rs("amtpaid1") + rs("amtpaid2")
				 
				 %>
                 <tr>
                 <td class="style29" style="width: 33%">
				 &nbsp;</td>
                 <td width="33%" class="style29" style="width: 33%">
				 &nbsp;</td>
                 <td width="33%" align="right" class="style15" style="width: 33%">
				 Total: $<%=formatnumber(otherpayments)%>
				 &nbsp;</td>
				 </tr>

                 <tr>
                 <td onclick="showPayment()" width="50%" style="cursor:pointer;background-image:url('newimages/rosubheader.jpg'); border:thin black outset; width: 100%;" class="style26" colspan="3">
				 <br><strong>Add/Delete Payments (Click)&nbsp;
                 </strong>&nbsp;<br><br></td>
                </tr>
                <%
                else
					paystmt = "select sum(`amt`) as payments from accountpayments where roid = " & request("roid") & " and shopid = '" & shopid & "'"
					set payrs = con.execute(paystmt)
					if isnull(payrs("payments")) then
					otherpayments = 0
					else
					otherpayments = payrs("payments")
					end if
					'response.write formatdollar(otherpayments)
					set payrs = nothing
					'response.write otherpayments & ":" & rs("amtpaid1") & ":" & rs("amtpaid2")
					customerpayments = otherpayments '+ rs("amtpaid1") + rs("amtpaid2")
                
                %>
                 <tr>
                 <td class="style29" style="width: 33%">
				 &nbsp;</td>
                 <td width="33%" class="style29" style="width: 33%">
				 &nbsp;</td>
                 <td width="33%" align="right" class="style15" style="width: 33%">
				 Total: $<%=formatnumber(otherpayments)%>
				 &nbsp;</td>
				 </tr>
                 <tr>
                 <td onclick="showPayment('swipe')" style="cursor:pointer;background-image:url('newimages/rosubheader.jpg'); border:thin black outset; width: 33%;" class="style26">
				 <br>Swipe Credit Card&nbsp;<br><br></td>
                 <td onclick="showPayment('echeck')" width="50%" style="cursor:pointer;background-image:url('newimages/rosubheader.jpg'); border:thin black outset; width: 33%;" class="style26">
				 Process E-Check</td>
                 <td onclick="showPayment('cash')" width="50%" style="cursor:pointer;background-image:url('newimages/rosubheader.jpg'); border:thin black outset; width: 33%;" class="style26">
				 Other Forms of Payment</td>
                </tr>
				<%

				end if
				%>
				
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
if taxparts > 0.01 or taxparts < 0 then
	salestax = round(taxparts * (rs("TaxRate") / 100),2)
else
	salestax = 0.00
end if
if ttllbr > 0 then
	labortax = round(ttllbr * (rs("LaborTaxRate") / 100),2)
else
	labortax = 0
end if

if ttlsub > 0 then
	subtax = round(ttlsub * (rs("SubletTaxRate") / 100),2)
else
	subtax = 0
end if

'check to see if we tax the userfee1
if userfee1taxable = "Taxable" then
	userfee1tax = round(userfee1amount * (rs("TaxRate") / 100),2) 
else
	userfee1tax = 0
end if

'response.write userfee1tax
nttltax = salestax + labortax + subtax + userfee1tax
		
xlist = "1734,1876,1888"
lar = split(xlist,",")
for j = lbound(lar) to ubound(lar)
	if lar(j) = request.cookies("shopid") then	
		if len(rs("userfee1")) > 0 then
			xuserfee1 = cdbl(rs("userfee1"))
		else
			xuserfee1 = 0
		end if
		if len(rs("userfee2")) > 0 then
			xuserfee2 = cdbl(rs("userfee2"))
		else
			xuserfee2 = 0
		end if
		if len(rs("userfee3")) > 0 then
			xuserfee3 = cdbl(rs("userfee3"))
		else
			xuserfee3 = 0
		end if
		
		xtotalfees = xuserfee1 + xuserfee2 + xuserfee3 + rs("hazardouswaste")
		'response.write xtotalfees
		xtotalfeetax = xtotalfees * (rs("TaxRate") / 100)
		
		nttltax = salestax + labortax + subtax + xtotalfeetax
		nttltax = round(nttltax,2)
		'response.write nttltax
	end if
next

ttlprts = notaxparts + taxparts

' ******************  DISCOUNT TAX ***************************
if rs("discounttaxable") = "no" then
	' response.write "notax"
elseif rs("discounttaxable") = "yes" then
	' take the discount tax amount off the tax amount
	discounttax = rs("discountamt") * (rs("TaxRate") / 100)
	nttltax = nttltax - discounttax
end if


             %>
             <tr>
              <td width="50%" align="right" class="style21" style="height: 20px">
			  Total Parts</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15" style="height: 20px"><%=formatdollar(ttlprts)%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" style="height: 22px" class="style21">
			  Total Labor</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" style="height: 22px" class="style15"><%=formatdollar(ttllbr)%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Total Sublet</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(ttlsub)%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td onclick="tempClose()" width="50%" align="right" class="style21">
			  Total Fees</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(rs("TotalFees"))%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Subtotal</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub)%></b>
			  &nbsp;&nbsp;
              <input type="hidden" name="discsubtotal" value="<%=cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub%>">
              </td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">
              <%
              if rs("TaxRate") = 0 and rs("labortaxrate") = 0 and rs("sublettaxrate") = 0 then
              	response.write "<span style='background-color:white;color:red;font-weight:bold;padding-right:10px;padding-left:10px;padding-top:2px;padding-bottom:2px;'>TAX EXEMPT</span>"
              end if
              %>
              Tax</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(nttltax)%></b>
			  &nbsp;&nbsp;</td>
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
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+nttltax-cdbl(rs("DiscountAmt")))%></b>
			  &nbsp;&nbsp;
              <input type="hidden" id="totalwithouttax" value="<%=cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub-cdbl(rs("DiscountAmt"))%>">
              <input type="hidden" id="discountforgp" value="<%=cdbl(rs("DiscountAmt"))%>">
			  </td>
             </tr>
             <tr>
              <td style="" width="50%" align="right" class="style21">Customer Payments</td>
              <td width="50%" bgcolor="#FFFFFF" style="color:red" align="right" class="style15">
			  &lt; <%=formatdollar(customerpayments)%> &gt;</b>&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Balance</td>
              <td width="50%" bgcolor="#FFFFFF" style="color:black" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+nttltax-cdbl(rs("DiscountAmt"))-customerpayments)%></b>
			  &nbsp;&nbsp;</td>
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
  balanceowed = formatnumber(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+nttltax-cdbl(rs("DiscountAmt"))-customerpayments)
'set rs = nothing
'strsql = "Select * from repairorders where shopid = '" & shopid & "' and ROID = " & request("roid")
'set nrs = con.execute(strsql)
for fnum = 0 to rs.fields.count-1
	if rs.fields(fnum).name = "ROID" then
		fldval = "<input type=hidden name='" & rs.fields(fnum).name
	else
		fldval = "<input type=hidden name='db" & rs.fields(fnum).name
	end if
	if rs.fields(fnum).name = "rodisc" then
		'do nothing
	elseif rs.fields(fnum).name = "warrdisc" then
		'do nothing
	elseif rs.fields(fnum).name = "TotalLbr" then
		response.write fldval & "' value='" & ttllbr & "'>" & chr(10)
	elseif rs.fields(fnum).name = "TotalPrts" then
		response.write fldval & "' value='" & ttlprts & "'>" & chr(10)
	elseif rs.fields(fnum).name = "TotalSublet" then
		response.write fldval & "' value='" & ttlsub & "'>" & chr(10)
	elseif rs.fields(fnum).name = "SalesTax" then
		response.write fldval & "' value='" & nttltax & "'>" & chr(10)
	elseif rs.fields(fnum).name = "TotalRO" then
		response.write fldval & "' value='" & ttllbr+ttlprts+ttlsub+nttltax & "'>" & chr(10)
	elseif rs.fields(fnum).name = "Subtotal" then
		response.write fldval & "' value='" & ttllbr+ttlprts+ttlsub & "'>" & chr(10)
	else
		response.write fldval & "' value='" & rs(fnum) & "'>" & chr(10)
	end if
	response.write "<input type=hidden name='fldtype" & rs.fields(fnum).name
	response.write "' value='" & rs.fields(fnum).type & "'>" & chr(10)
next

'calc gp
ttls = ttlprts+ttllbr
ttlc = pc+tlc
dprofit = ttls-ttlc
'if dprofit > 0 then gp = dprofit/ttls
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
<div id="cctimer">
	Processing...<br><br>
<img src="newimages/ajax-loader.gif">
</div>

<div id="estimator"><img style="position:absolute;top:40%;left:45%" src="newimages/ajax-loader.gif"></div>
<form method="post" name="guideform" action="addguidecontents.asp">
<input type="hidden" name="laboritems"><input type="hidden" name="partitems"><input type="hidden" name="comid"><input type="hidden" name="roid" value="<%=request("roid")%>"><input type="hidden" name="shopid" value="<%=shopid%>">
<div id="estimatorqueue"><b>Labor To be added to RO:</b>  <span style="font-size:x-small">
	(make any changes you like to the values from the guide, then click on Add 
	All)</span><br></div>
<div id="pestimatorqueue"><b>Parts To be added to RO:</b>  <span style="font-size:x-small">
	(make any changes you like to the values from the guide, then click on Add 
	All)</span><br></div>
<div id="techselect">
	Please select a technician for this labor operation
<select size="1" name="tech">

<%
strsql = "Select EmployeeLast, EmployeeFirst from employees where shopid = '" & shopid & "' and Active = 'yes'"
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
<div style="display:none;width:500px;" id="popup">
<form name="pmtform" method="post">
<%
if merchantaccount = "no" then
%>
<input type="hidden" name="roid" value="<%=roid%>">
<input type="hidden" name="cid" value="<%=cid%>">
<input type="hidden" name="swipestring" value="">
<table style="width: 100%" align="center" class="style1" cellspacing="0" cellpadding="3">
	<tr>
		<td style="width: 50%"><strong>Amount:&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input value="<%=replace(balanceowed,",","")%>" id="amt" name="amt" type="text" ></td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>How Paid:&nbsp; </strong></td>
		<td style="width: 50%" class="style3">
		<select name="type">
                   <%
					stmt = "select `method` from paymentmethods where shopid = '" & shopid & "'"
					set pmrs = con.execute(stmt)
					if not pmrs.eof then
						do until pmrs.eof
							response.write "<option value='" & pmrs("method") & "'>" & pmrs("method") & "</option>"
							pmrs.movenext
						loop
					end if

                   %>
		</select>
		</td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>Number (optional):&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input name="number" type="text" /></td>
	</tr>
	<tr>
		<td colspan="2" class="style9">
		<input type="button" onclick="savePayment()" name="b" value="Post Payment">
		<input name="Button1" type="button" onclick="closePayment()" value="Cancel">
		</td>
	</tr>
</table>
<%
else
%>
<table style="display:none" id="othertable" style="width: 100%" align="center" class="style1" cellspacing="0" cellpadding="3">
	<tr>
		<td style="width: 50%"><strong>Amount:&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input value="<%=replace(balanceowed,",","")%>" id="cashamt" name="amt" type="text" ></td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>How Paid:&nbsp; </strong></td>
		<td style="width: 50%" class="style3">
		<select name="cashtype">
                   <%
					stmt = "select `method` from paymentmethods where shopid = '" & shopid & "'"
					set pmrs = con.execute(stmt)
					if not pmrs.eof then
						do until pmrs.eof
							response.write "<option value='" & pmrs("method") & "'>" & pmrs("method") & "</option>"
							pmrs.movenext
						loop
					end if

                   %>
		</select>
		</td>
	</tr>
	<tr>
		<td style="width: 50%"><strong>Number (optional):&nbsp; </strong></td>
		<td style="width: 50%" class="style3"><input name="cashnumber" type="text" /></td>
	</tr>
	<tr>
		<td colspan="2" class="style9">
		<input type="button" onclick="savePayment()" name="b" value="Post Payment">
		<input name="Button1" type="button" onclick="closePayment()" value="Cancel">
		</td>
	</tr>
</table>

<input type="hidden" name="roid" value="<%=roid%>">
<input type="hidden" name="cid" value="<%=cid%>">
<input type="hidden" name="cardstring">
<input type="hidden" name="processtype" value="cc">
<span style="font-size:12px;">
Enter any form of payment you wish here.  It will be recorded in Shop Boss Pro<br><br>
<b>*** NOTE ***</b> ANY CREDIT CARD INFORMATION ENTERED HERE WILL <b>NOT BE AUTOMATICALLY PROCESSED</b>.  YOU WILL NEED TO CLICK ON <b>SWIPE CREDIT CARD</b> OR PROCESS THE CREDIT CARD MANUALLY
</span>	


<table id="cctable" cellpadding="3" cellspacing="0" style="width: 100%;display:none">
		<tr>
			<td class="style31">Swipe Credit Card</td>
			<td>
			<input onkeypress="return noenter(event)" onfocus="this.style.backgroundColor='yellow'" onblur="this.style.backgroundColor='white'" id="cardswipe" name="cardswipe" type="password"> 
			<span onclick="clearSwipe()" style="color:maroon;font-weight:bold;cursor:pointer">
			Clear</span></td>
		</tr>
		<tr>
			<td class="style32">Amount</td>
			<td><input name="ccamt" value="<%=replace(balanceowed,",","")%>" type="text" >&nbsp;</td>
		</tr>
		<tr>
			<td class="style32">Card Number</td>
			<td><input name="cc" type="text" >&nbsp;</td>
		</tr>
		<tr>
			<td class="style32">Expiration Date</td>
			<td>
			<select style="width:50px;" name="expmonth">
				<option value="01">01</option>
				<option value="02">02</option>
				<option value="03">03</option>
				<option value="04">04</option>
				<option value="05">05</option>
				<option value="06">06</option>
				<option value="07">07</option>
				<option value="08">08</option>
				<option value="09">09</option>
				<option value="10">10</option>
				<option value="11">11</option>
				<option value="12">12</option>
			</select> &nbsp;&nbsp;
			<select style="width:75px;" name="expyear">
				<%
				snum = datepart("yyyy",date)
				snum = cint(snum)
				for j = snum to snum+10
					twodigityear = j - 2000
				%>
				<option value="<%=twodigityear%>"><%=j%></option>
				<%
				next
				%>
			</select>
			</td>
		</tr>
		<tr>
			<td class="style32">Payment Type</td>
			<td>
			<select name="cctype">
	           <%
				stmt = "select `method` from paymentmethods where shopid = '" & shopid & "'"
				set pmrs = con.execute(stmt)
				if not pmrs.eof then
					do until pmrs.eof
						response.write "<option value='" & pmrs("method") & "'>" & pmrs("method") & "</option>"
						pmrs.movenext
					loop
				end if
	
	           %>
			</select>

			</td>
		</tr>
		<tr>
			<td class="style32">Reference Number</td>
			<td>
			<input name="ccnumber" type="text"></td>
		</tr>
		<tr>
			<td class="style32">First Name on Card</td>
			<td>
			<input name="ccfirstname" type="text"></td>
		</tr>
		<tr>
			<td class="style32">Last Name on Card</td>
			<td>





<font color="#FFFFFF">


			<input name="cclastname" type="text"></td>
		</tr>
		<tr>
			<td colspan="2" class="style33">
				<input type="button" onclick="postCCPayment()" name="b" value="Post Payment">
				<input name="Button1" type="button" onclick="closePayment()" value="Cancel">

			</td>
		</tr>
	</table>
<table id="achtable" cellpadding="3" cellspacing="0" style="width: 100%; display:none">
		<tr>
			<td class="style32">Amount</td>
			<td><input name="achamt" value="<%=replace(balanceowed,",","")%>" type="text" >&nbsp;</td>
		</tr>
		<tr>
			<td class="style32">Checking Account #</td>
			<td>
			<font color="#FFFFFF"><input name="accountnumber" type="text" ></td>
		</tr>
		<tr>
			<td class="style32">Routing Number</td>
			<td>
			<font color="#FFFFFF"><input name="routing" type="text" ></td>
		</tr>
		<tr>
			<td class="style32">First Name on Account</td>
			<td>
			<input name="achfirstname" type="text"></td>
		</tr>
		<tr>
			<td class="style32">Last Name on Account</td>
			<td>





<font color="#FFFFFF">


			<input name="achlastname" type="text"></td>
		</tr>
		<tr>
			<td colspan="2" class="style33">
				<input type="button" onclick="postCCPayment()" name="b" value="Post Payment">
				<input name="Button1" type="button" onclick="closePayment()" value="Cancel">

			</td>
		</tr>
	</table>


<%
end if
%>
</form>
<br>
	Payments Posted to this RO

	<table cellpadding="2" cellspacing="0" style="width: 100%">
		<tr>
			<td style="background-color:maroon;color:white;font-weight:bold">
			Amount&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">
			Type&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">
			Date&nbsp;</td>
			<td style="background-color:maroon;color:white;font-weight:bold">
			Delete&nbsp;</td>
		</tr>
		<%
		stmt = "select * from accountpayments where roid = " & request("roid") & " and shopid = '" & shopid & "'"
		set payrs = con.execute(stmt)
		if not payrs.eof then
			do until payrs.eof
		%>
		<tr>
			<td><%=formatnumber(payrs("amt"))%>&nbsp;</td>
			<td><%=payrs("ptype")%>&nbsp;</td>
			<td><%=payrs("pdate")%>&nbsp;</td>
			<%
			if request.cookies("deletepayments") = "yes" or request.cookies("deletepayments") = "" then
			%>
			<td onclick="deletePayment('<%=payrs("id")%>')" style="color:#336699;cursor:pointer">
			Delete&nbsp;</td>
			<%
			else
			%>
			<td >&nbsp;</td>
			<%
			end if
			%>
		</tr>
		<%
				payrs.movenext
			loop
		else
			response.write "<tr><td colspan='4'>No other payments found</td></tr>"
		end if
		%>
	</table>



</div>
<div id="popupdropshadow"></div>

	<div style="padding:20px;padding-left:40px;width:400px;height:300px;position:absolute;left:45%;top:40%;z-index:999;display:none;background-color:white;border:medium black outset;color:black" id="techfortime<%'crs("laborid")%>">
	<br><br>
	<form name="techfortimeform">
	<input type="hidden" name="roid" value="<%=request("roid")%>">
	<input type="hidden" name="comid">
	<input type="hidden" name="labortimeclockid">
	<input type="hidden" name="shopid" value="<%=shopid%>">
	<input type="hidden" name="laborid"><input type="hidden" name="startdate"><input type="hidden" name="enddate">

		Select the technician for this time:<br>
	<select name="techformtimetech">
		<option selected="selected"></option>
		<%
		stmt = "select * from employees where shopid = '" & shopid & "' and lcase(active) = 'yes'"
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
		max=999999
		min=111111
		Randomize
		r = (Int((max-min+1)*Rnd+min))

		%>
	</select><br><br>
	<input onclick="postLaborTimeClock()" name="Button1" type="button" value="Post Time" > <input onclick="cancelTimeClock()" name="Button1" type="button" value="Cancel" >
	</form>
	</div>


<iframe src="rohistorydata.asp?<%if origshopid <> currshopid and len(origshopid) > 0 then response.write "r=" & r & "&origshopid=" & origshopid & "&"%>roid=<%=request("roid")%>&vid=<%=rs("vehid")%>&vin=<%=rs("vin")%>" style="display:none" id="history"></iframe>
<input style="display:none" type="button" onclick="closeInspection()" id="inspectionbutton" value="Close">
<iframe src="" style="display:none" id="inspection"></iframe>
<iframe frameborder="0" src="upload/upload.asp?roid=<%=request("roid")%>" style="border:medium black outset;width:60%;height:80%;position:absolute;left:20%;top:10%;display:none;z-index:999" id="showframe"></iframe>
<iframe id="picframe" frameborder="0" src='upload/showpics.asp?roid=<%=request("roid")%>' style="border:medium black outset;width:90%;height:90%;position:absolute;left:5%;top:5%;display:none;z-index:999"></iframe>

<div style="border:2px black solid;color:black;font-family:arial;text-align:center;z-index:999;display:none;position:absolute;top:10%;left:20%;width:60%;height:40%;background-color:white" id="alerts">
	<br><strong>You have changed the status of this Repair Order.</strong><br><br>
	Would you like to notify the customer via text message or email?

<%
'check for email and cell number and provider

'get email from customer record
stmt = "select * from customer where customerid = " & rs("customerid") & " and shopid = '" & shopid & "'"
set ers = con.execute(stmt)
email = ers("email")
set ers = nothing

if len(rs("cellphone")) > 0 then
	'we have a cell number, do we have a provider?
	
	if len(rs("cellprovider")) > 0 then
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
<div style="text-align:center;border:15px black solid;color:red;background-color:white;width:90%;height:90%;position:absolute;top:5%;left:5%;z-index:1050;display:none" id="invoicediv">
	<strong>Your Invoice is displayed below.</strong>
<input onclick="closePrint()" name="Button1" type="button" value="Close" >
<%
if esign = "yes" then
%>
<input onclick="esign()" name="Button1" type="button" value="E-sign (IPad)" >
<input onclick="sigPad()"  name="Button1" type="button" value="Scriptel Signature Pad" >
<%
end if
%>
<input onclick="showEmailAddress('no')" name="Button1" type="button" value="E-mail" >
<%
if esign = "yes" then
%>
<input onclick="printIframe('invoice')" id="printbutton" value="Print" type="button">
<%
end if
%>
<iframe frameborder="0" scrolling="no" style="width:95%;height:95%;z-index:1599" id="invoice" name="invoice"></iframe>

</div>
<script type="text/javascript" language="javascript">

function printIframe(id)
{
	
    var pdf = document.getElementById("invoice").src
    //document.getElementById("temp").innerHTML = pdf
    pdf = pdf.replace("http://<%=request.servervariables("SERVER_NAME")%>/sbp/","")
    pdf = pdf.replace("https://<%=request.servervariables("SERVER_NAME")%>/sbp/","")
    document.getElementById("invoice").src = ""
    //document.getElementById("hider").style.display = "none"
    //alert(pdf)
    xmlHttp=GetXmlHttpObject();
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	}
    xmlHttp.onreadystatechange =
        function(){
            if (xmlHttp.readyState == 4 && xmlHttp.status == 200){
            	//alert(xmlHttp.responseText)
            	//document.getElementById("hider").style.display = "none"
            	rt = xmlHttp.responseText
	             document.getElementById("invoice").src = rt;

            }
        };
	
    var url = "roprintpdf2.asp?roid=<%=request("roid")%>&shopid=<%=request.cookies("shopid")%>&pdf="+pdf;
    //alert(url)
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}
function esign(){

	var x = (window.innerWidth - 700) / 2
	document.getElementById("invoice").src = ""
	document.getElementById("invoice").style.display = "none"
	document.getElementById("invoicediv").style.display = "none";
	
	document.getElementById("signature").style.left = x+"px"
	document.getElementById("signature").style.display = "block"
	document.getElementById("signature").src = "signature.asp?shopid=<%=request("shopid")%>&roid=<%=request("roid")%>"

}

function sigPad(){
	var x = (window.innerWidth - 700) / 2
	document.getElementById("invoice").src = ""
	document.getElementById("invoice").style.display = "none"
	document.getElementById("invoicediv").style.display = "none";
	
	document.getElementById("signature").style.left = x+"px"
	document.getElementById("signature").style.display = "block"
	document.getElementById("signature").src = "signaturepad.asp?shopid=<%=request("shopid")%>&roid=<%=request("roid")%>"
	document.getElementById("signature").focus()
	document.getElementById("signature").contentWindow.focus()


}
function closePrint(){


	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("hider").style.display = "none"
	document.getElementById("invoice").src = ""
	document.getElementById("invoice").style.display = "none"
	if (document.getElementById("printedrolist")){
		// reload the list with ajax call
        xmlHttp=GetXmlHttpObject();
		if (xmlHttp==null){
			alert ("Browser does not support HTTP Request")
			return
		}
        xmlHttp.onreadystatechange =
            function(){
                if (xmlHttp.readyState == 4 && xmlHttp.status == 200){
                    document.getElementById("printedrolist").innerHTML = xmlHttp.responseText;
                    //alert(xmlHttp.responseText)
                }
            };

        var url = "roreloadprintedinvoices.asp?roid=<%=request("roid")%>&shopid=<%=request.cookies("shopid")%>";
		xmlHttp.open("GET",url,true)
		xmlHttp.send(null)
	}

	if(document.getElementById("printdrop").style.display == "inline-block"){
		document.getElementById("hider").style.display = "block"
	}
}		

<%
if request("printro") = "yes" or request("printwo") = "yes" then
	max=999999
	min=111111
	Randomize
	r = Int((max-min+1)*Rnd+min)

%>


x = setTimeout("printRO()",200)

function printRO(){
	//alert("printing")
	document.getElementById("hider").style.display = "block"
	document.getElementById('timer').style.display = 'block'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	<%
	if request("printro") = "yes" then
		if request("test") = "test" then
	%>
			var url="pdfinvoices/printpdfro.asp?shopid=1238&r=<%=r%>&test=test&roid=<%=request("roid")%>"
	<%
		else
	%>
			var url="pdfinvoices/printpdfro.asp?shopid=1238&r=<%=r%>&roid=<%=request("roid")%>"
	<%
		end if
	elseif request("printwo") = "yes" then
	%>
		var url="pdfinvoices/printpdfwo.asp?shopid=1238&r=<%=r%>&roid=<%=request("roid")%>"
	<%
	end if
	%>
	//alert(url)
	xmlHttp.onreadystatechange=printROStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}


function printROStateChanged(){
	document.getElementById("hider").style.display = "block"
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){

		//alert(xmlHttp.responseText)
		invname = xmlHttp.responseText
		//document.getElementById("error").innerHTML = invname
		<%
		if printpopup = "no" then
		%>
		document.getElementById("hider").style.display = "block"
		document.getElementById("invoicediv").style.display = "block"
		document.getElementById("invoice").src = "pdfinvoices"+invname
		document.getElementById('timer').style.display = 'none'
		
		document.getElementById("pdf").value = "pdfinvoices"+invname
		//alert("pdf:"+document.getElementById("pdf").value)
		
		//document.getElementById("invoice").focus()
		//document.getElementById("invoice").contentWindow.print()

		<%
		elseif printpopup = "yes" then
		%>
		//document.getElementById('timer').style.display = 'none'
		//document.getElementById("hider").style.display = "block"
		//document.getElementById("ipadprint").style.display = "block"
		//document.getElementById("printpdfbutton").style.display = "block"
		//location.href="ropopuppdf.asp?roid=<%=request("roid")%>&fname="+"pdfinvoices"+invname
		//document.getElementById("ipadpdflocation").value = "ropopuppdf.asp?roid=<%=request("roid")%>&fname="+"pdfinvoices"+invname
		document.getElementById("printpdfbutton").href = "ropopuppdf.asp?roid=<%=request("roid")%>&fname="+"pdfinvoices"+invname
		//document.getElementById("ipadprint").onclick = printPopUp("ropopuppdf.asp?roid=<%=request("roid")%>&fname="+"pdfinvoices"+invname,"mywin")
		document.getElementById("hider").style.display = "block"
		document.getElementById("invoicediv").style.display = "block"
		document.getElementById("printpdfbutton").style.display = "block"
		document.getElementById("invoice").src = "pdfinvoices"+invname
		document.getElementById('timer').style.display = 'none'
		document.getElementById("pdf").value = "pdfinvoices"+invname
		//alert("pdf:"+document.getElementById("pdf").value)

		//document.getElementById("invoice").focus()
		//document.getElementById("invoice").contentWindow.print()


		<%
		end if
		%>
	}
}

function printInvoicePDF(){
}
<%
end if
if request("emro") = "yes" then
%>
x = setTimeout("emailRO()",1000)

function emailRO(){
	//alert("printing")
	document.getElementById("hider").style.display = "block"
	document.getElementById('timer').style.display = 'block'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
<%
l = "1471"
lar = split(l,",")
for j = lbound(lar) to ubound(lar)
	if request.cookies("shopid") = lar(j) and rostat = "1INSPECTION" then
		response.write "var url = 'pdfinvoices/" & lar(j) & "/pdffloridaestimate.asp?roid=" & request("roid") & "&shopid=" & request.cookies("shopid") & "'"
	else
%>
	var url="pdfinvoices/printpdfro.asp?shopid=1238&r=<%=r%>&roid=<%=request("roid")%>"
<%
	end if
next
%>
	//alert(url)
	xmlHttp.onreadystatechange=emailROStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}


function emailROStateChanged(){
	document.getElementById("hider").style.display = "block"
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){

		//alert(xmlHttp.responseText)
		invname = xmlHttp.responseText
		//alert(invname)
		sendEmailRO(invname)
		
	}
}
<%
if len(request.querystring("email")) > 0 then
	customeremail = request.querystring("email")
end if
if request.querystring("updateemail") = "yes" then
%>
document.theform.dbemail.value = '<%=request.querystring("email")%>'<%=chr(10)%>
<%
' now update the customer record with new email
stmt = "update customer set email = '" & request.querystring("email") & "' where shopid = '" & request.cookies("shopid") & "' and customerid = " & cid
'response.write stmt
con.execute(stmt)
end if
%>

function sendEmailRO(f){
	//alert("printing")
	document.getElementById("hider").style.display = "block"
	document.getElementById('timer').style.display = 'block'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="pdfinvoices/sendinvoice.asp?r=<%=r%>&shopid=<%=request.cookies("shopid")%>&invpath="+f+'&shopname=<%=server.urlencode(replace(shopname,"&"," and "))%>&email=<%=customeremail%>'
	//alert(url)
	xmlHttp.onreadystatechange=sendEmailROStateChanged
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function sendEmailROStateChanged(){document.getElementById("hider").style.display="block",(4==xmlHttp.readyState||"complete"==xmlHttp.readyState)&&(rt=xmlHttp.responseText,alert("sent"!=rt?"There was an error sending the RO to the customer email.  Please verify their email address":"Invoice Sent"),document.getElementById("hider").style.display="none",document.getElementById("timer").style.display="none")}
<%
end if
if merchantaccount = "no" then
%>

function focusPayment(){

	document.pmtform.amt.focus()
}
<%
else
%>
document.pmtform.cardswipe.focus()
function validSwipe(){return cardswipe=document.pmtform.cardswipe.value,lastchar=cardswipe.indexOf("E?"),lastchar?(alert("Bad card read.  Please try again"),document.pmtform.cardswipe.value="",void document.pmtform.cardswipe.focus()):void 0}function noenter(e){e=e||window.event;var d=e.keyCode||e.charCode;return 13==d?(document.pmtform.cc.disabled="disabled",document.pmtform.expmonth.disabled="disabled",document.pmtform.expyear.disabled="disabled",document.pmtform.ccfirstname.disabled="disabled",document.pmtform.cclastname.disabled="disabled",document.pmtform.ccnumber.disabled="disabled",ccnum=SwipeParser(document.pmtform.cardswipe.value),13!==d):(document.pmtform.cc.disabled=!1,document.pmtform.expmonth.disabled=!1,document.pmtform.expyear.disabled=!1,document.pmtform.ccfirstname.disabled=!1,document.pmtform.cclastname.disabled=!1,document.pmtform.ccnumber.disabled=!1,void 0)}function changeACH(){"Process E-Check"==document.getElementById("ach").innerText?(document.getElementById("ach").innerText="Process Credit Card",document.pmtform.processtype.value="ach",document.getElementById("cctable").style.display="none",document.getElementById("achtable").style.display="",document.pmtform.accountnumber.focus()):(document.getElementById("ach").innerText="Process E-Check",document.pmtform.processtype.value="cc",document.getElementById("cctable").style.display="",document.getElementById("achtable").style.display="none",document.pmtform.cardswipe.focus())}function clearSwipe(){document.pmtform.cardswipe.value="",document.pmtform.cc.disabled=!1,document.pmtform.expmonth.disabled=!1,document.pmtform.expyear.disabled=!1,document.pmtform.ccfirstname.disabled=!1,document.pmtform.cclastname.disabled=!1,document.pmtform.ccnumber.disabled=!1,document.pmtform.cardswipe.focus()}<%
end if
%>

function focusPayment(){document.pmtform.cardswipe.focus()}function SwipeParser(r){if(blnCarrotPresent=!1,blnEqualPresent=!1,blnBothPresent=!1,strCarrotPresent=r.indexOf("^"),strEqualPresent=r.indexOf("="),strCarrotPresent>0&&(blnCarrotPresent=!0),strEqualPresent>0&&(blnEqualPresent=!0),1==blnEqualPresent&&1==blnCarrotPresent)r=""+r+" ",arrSwipe=new Array(4),arrSwipe=r.split("^"),arrSwipe.length>2?(account=stripAlpha(arrSwipe[0].substring(1,arrSwipe[0].length)),firstccnum=account.substring(0,1),"5"==firstccnum?document.pmtform.cctype.value="Mastercard":"4"==firstccnum?document.pmtform.cctype.value="Visa":"3"==firstccnum?document.pmtform.cctype.value="American Express":"6"==firstccnum&&(document.pmtform.cctype.value="Discover")):alert("Error Parsing Card.  \r\n Please Contact MIS/IT! \r\n");else if(1==blnCarrotPresent)r=""+r+" ",arrSwipe=new Array(4),arrSwipe=r.split("^"),arrSwipe.length>2?(account=stripAlpha(arrSwipe[0].substring(1,arrSwipe[0].length)),firstccnum=account.substring(0,1),"5"==firstccnum?document.pmtform.cctype.value="Mastercard":"4"==firstccnum?document.pmtform.cctype.value="Visa":"3"==firstccnum?document.pmtform.cctype.value="American Express":"6"==firstccnum&&(document.pmtform.cctype.value="Discover")):alert("Error Parsing Card.  \r\n Please Contact MIS/IT! \r\n");else{if(1!=blnEqualPresent)return validSwipe();sCardNumber=r.substring(1,strEqualPresent),sYear=r.substr(strEqualPresent+1,2),sMonth=r.substr(strEqualPresent+3,2),account=sAccountNumber=stripAlpha(sCardNumber),firstccnum=account.substring(0,1),"5"==firstccnum?document.pmtform.cctype.value="Mastercard":"4"==firstccnum?document.pmtform.cctype.value="Visa":"3"==firstccnum?document.pmtform.cctype.value="American Express":"6"==firstccnum&&(document.pmtform.cctype.value="Discover")}}function stripAlpha(r){return null==r?"":r.replace(/[^0-9]/g,"")}function selectOptionByValue(r,e){for(var t=r.options,c=t.length;c;)t[--c].value==e&&(r.selectedIndex=c,c=0)}
</script>
<div id="error"></div>
<input id="ipadprint" style="width:300px;height:100px;font-weight:bold;color:red;display:none;position:absolute;left:10px;top:10px;z-index:999" onclick="printPopUp()" name="Button1" type="button" value="RO Generated.  Click to Open/Print" />
<input id="ipadpdflocation" name="Hidden1" type="hidden" />
<div id="temp"></div>
<script language="javascript">
status = "Shop Boss Pro"

function goAnchor(){

	n = getCookie("anchor")
	//alert(n)
	window.location.hash = "#"+n
	del_cookie()

}
x = setTimeout("goAnchor()",500)
function del_cookie()
{
    document.cookie = 'anchor=; expires=Thu, 01 Jan 1970 00:00:01 GMT;';
}
function alertChangeOrder(){
	document.theform.changeorder.value = "yes"
}

</script>
<!-- #include file=aspscripts/im.asp -->
</html>

<%
'Copyright 2011 - Boss Software Inc.
%>
