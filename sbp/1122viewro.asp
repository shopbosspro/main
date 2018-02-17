<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=aspscripts/salesTax.asp -->
<%
shopid = request.cookies("shopid")
set fso = server.createobject("scripting.filesystemobject")
f = fso.getfolder(server.mappath("."))
if fso.fileexists(f & "\" & request.cookies("shopid") & "viewro.asp") then
	'response.redirect request.cookies("shopid") & "viewro.asp?roid=" & request("roid")
end if
set fso = nothing

'get esign
stmt = "select esign from company where shopid = '" & shopid & "'"
set rs = con.execute(stmt)
esign = rs("esign")

shopid = request.cookies("shopid")

' check sales tax
checkSalesTax request("roid"),request.cookies("shopid")


%>

<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<style type="text/css">
<!--
a            { text-decoration: none }
-->
select{
	width:150px;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<%
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
myvid = rs("vehid")
customerid = rs("customerid")
cemail = rs("email")
'calculate gp
if len(rs("tagnumber")) > 0 then
	roid = rs("tagnumber")
else
	roid = rs("roid")
end if
pc = rs("partscost")
if rs("datein") >= #11/11/2005# then
	pro = "printronewall.asp?sp=viewro.asp"
else
	pro = "printro.asp"
end if

homephone = "(" & left(rs("CustomerPhone"),3) & ") " & mid(rs("CustomerPhone"),4,3) & "-" & right(rs("CustomerPhone"),4)
workphone = "(" & left(rs("CustomerWork"),3) & ") " & mid(rs("CustomerWork"),4,3) & "-" & right(rs("CustomerWork"),4)
cellphone = "(" & left(rs("CellPhone"),3) & ") " & mid(rs("CellPhone"),4,3) & "-" & right(rs("CellPhone"),4)
'ttlsub = rs("totalsublet")
'response.write ttlsub
ttlprts = rs("totalprts")
'ttllbr = rs("totallbr")
if ucase(rs("Status")) = "CLOSED" then 
	if request("sp") = "history.asp" then
		'newsp = "history.asp?lic=" & request("lic") & "&roid=" & request("roid") & "&sp=" & request("sp")
		newsp = "history.go(-1)"
	elseif request("sp") = "findroframe.asp" then
		newsp = "location.href='viewro.asp?roid=" & request("roid") & "&srch=" & request("srch") & "&sf=" & request("sf") & "&SortBy=" & request("SortBy") & "&sp=" & request("sp") & "'"
	elseif request("sp") = "wip.asp" then
		newsp = "location.href='viewro.asp?roid=" & request("roid") & "'"
	else
		newsp = "location.href='wip.asp'"
	end if
end if

%>

<script type="text/javascript" src="javascripts/ro.js"></script>
<style type="text/css">
<!--
p, td, th, li { font-size: 10pt; font-family: Verdana, Arial, Helvetica }
.style1 {
	border-width: 0;
	font-family: Arial;
	color: white;
	/*background-color: #0066CC;*/
}
.abuts{
	font-size:10px;
	width:50px;
	height:18px;
	cursor:hand;
}
#hider{
	position:absolute;
	top:0px;
	left:0px;
	width:100%;
	height:100%;
	background-color:gray;
	-ms-filter:"progid:DXImageTransform.Microsoft.Alpha(Opacity=70)";
	filter: alpha(opacity=50); 
	-moz-opacity:.50; 
	opacity: .5;
	z-index:998;
	display:none;

}
#timer{
	background-color:white;
	position:absolute;
	top:25%;
	left:25%;
	width:50%;
	height:20%;
	z-index:999;
	display:none;
	text-align:center;
	border:4px navy outset;
	color:black;
	padding-top:10%;
	font-family:Verdana, Arial, sans-serif;
	font-weight:bold;
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
	width:12.5%;
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

.style22 {
	color: #000000;
	border-width: 0;
	background-color: #FFFFFF;
}
.style23 {
	text-align: right;
	color: #FFFFFF;
	background-color: #0066CC;
}
.style24 {
	color: #FFFFFF;
	background-color: #0066CC;
}

</style>
<link href="jquery/jquery-ui.css" rel="stylesheet">
<link href="css/bootstrap.min.css" rel="stylesheet">
<script type="text/javascript" language="javascript" src="jquery/jquery-1.10.2.min.js"></script>
<script type="text/javascript" language="javascript" src="jquery/jquery-ui.js"></script>
<script type="text/javascript" language="javascript" src="javascripts/bootstrap.min.js"></script>
<script type="text/javascript" language="javascript">

function showDetails(partid){
	
	$.ajax({
		data: "partid="+partid+"&shopid=<%=shopid%>",
		url:  "aspscripts/getpartdetails.asp",
		success: function(r){
			console.log(r)
			if (r.indexOf("|") >= 0){
				rar = r.split("|")
				pn = rar[0]; pd = rar[1]; pc = rar[2]; pp = rar[3]; pq = rar[4]; ps = rar[5]; po = rar[6]; pi = rar[7]; pcat = rar[8];
				mb = "<b>Part Number:</b> " + pn + "<BR>"
				mb += "<b>Description:</b> " + pd + "<BR>"
				mb += "<b>Selling Price:</b> " + pp + "<BR>"
				mb += "<b>Cost:</b> " + pc + "<BR>"
				mb += "<b>Qty:</b> " + pq + "<BR>"
				mb += "<b>Ext. Price:</b> " + (pp * pq) + "<BR>"
				mb += "<b>PO #</b> " + po + "<BR>"
				mb += "<b>Supplier:</b> " + ps + "<BR>"
				mb += "<b>Invoice #</b> " + pi + "<BR>"
				mb += "<b>Category:</b> " + pcat + "<BR>"
				$('#modal-body').html(mb)
				$('#modal-title').html("Part Details")
			}else{
				$('#modal-body').html("That part was not found")
				$('#modal-title').html("Part Details")
			}
		}
	
	
	}),
	$('#myModal').modal('show');
	

}
function showPics(){

	document.getElementById("picframe").style.display = "block"
	document.getElementById("hider").style.display = "block"


}


var xmlHttp
function showHistory(){
	document.getElementById("hider").style.display = "block"
	document.getElementById("history").style.display = "block"

}


function showSavedROs(){

	document.getElementById("hider").style.display = "block"
	document.getElementById("printdrop").style.display = "block"

}

function reprintRO(n){

	filepath = "savedinvoices/<%=shopid%>/<%=request("roid")%>/"+n
	
	document.getElementById("invoice").src = filepath
	document.getElementById("invoice").style.width = "100%"
	document.getElementById("invoice").style.display = "block"
	document.getElementById("invoicediv").style.display = "block"
	//document.getElementById('invoice').reload(true)
}

function showEmailAddress(inclprint){

	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("invoice").style.display = "none"
	document.getElementById("sendemail").style.display = "block"
	document.getElementById("hider").style.display = "block"
	document.getElementById("inclprint").value = inclprint
	//document.getElementById('printdrop').style.display='none';

}

function closePrint(){
	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("invoice").src = ""
	document.getElementById("invoice").style.display = "none"

}

function printIframe(id)
{
	
    var pdf = document.getElementById("invoice").src
    document.getElementById("invoice").src = ""
    //alert(pdf)
    xmlHttp=new XMLHttpRequest()
    xmlHttp.onreadystatechange =
        function(){
            if (xmlHttp.readyState == 4 && xmlHttp.status == 200){
            	//alert(xmlHttp.responseText)
                document.getElementById("invoice").src = xmlHttp.responseText;
            }
        };

    var url = "roprintpdf.asp?pdf="+pdf;
    //alert(url)
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}

function redirectWithEmail(){

	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("invoice").style.display = "none"
	invmessage = encodeURIComponent(document.getElementById("emailinvoicemessage").value)
	invmessage = invmessage.replace(/&/g,"&amp")
	invmessage = invmessage.replace(/#/g,"%23")
	submessage = document.getElementById("messagesubject").value
	submessage = submessage.replace(/&/g,"&amp")
	submessage = submessage.replace(/#/g,"%23")
	//document.getElementById("hider").style.display = "none"
	var inclprint = document.getElementById("inclprint").value
	var curremail = document.getElementById("curremailaddress").value
	if(document.getElementById("updateemail").checked ==true){
		updateemail = "yes"
	}
	if(document.getElementById("updateemail").checked ==false){
		var updateemail = "no"
	}
	
	xmlHttp=new XMLHttpRequest()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	email = document.getElementById("curremailaddress").value
	if (document.getElementById("updateemail").checked ==true){update = "yes"}else{update="no"}
	var pdfpath = document.getElementById("invoice").src
	var url="roemailinvoice.asp?sl="+submessage+"&message="+invmessage+"&email="+email+"&update="+update+"&pdfpath="+pdfpath+"&shopid=<%=shopid%>"
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

function showReopenHistory(){

	xmlHttp=new XMLHttpRequest()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="reopenhistory.asp?roid=<%=request("roid")%>&shopid=<%=shopid%>"
	//alert(url)
	xmlHttp.onreadystatechange=function(){
			if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
				rt = xmlHttp.responseText
				document.getElementById("reopenhistory").innerHTML = rt
				document.getElementById("reopenhistory").style.display = "block"
				document.getElementById("hider").style.display="block"
			}
		}

	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)

}

function closeReopenHistory(){

	document.getElementById("reopenhistory").innerHTML = ""
	document.getElementById("reopenhistory").style.display = "none"
	document.getElementById("hider").style.display="none"


}
</script>
</head>

<body  link="#800000" vlink="#800000" alink="#800000" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<div id="myModal" class="modal fade" role="dialog">
  <div class="modal-dialog">

    <!-- Modal content-->
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal">×</button>
        <h4 id="modal-title" class="modal-title">Modal Header</h4>
      </div>
      <div id="modal-body" class="modal-body">
        <p>Some text in the modal.</p>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
      </div>
    </div>

  </div>
</div>
<div style="width:40%;position:absolute;height:300px;left:30%;top:200px;z-index:1000;display:none;background-color:white;border:2px black solid;overflow-y:scroll" id="reopenhistory"></div>
<input type="hidden" id="inclprint" name="inclprint" value="">
<div style="position:absolute;top:100px;width:400px;height:450px; left:40%;background-color:white;padding:25px;display:none;z-index:1005;border:2px black outset" id="sendemail">
	Please verify customer email. You can change it if necessary<br>
<input type="text" id="curremailaddress" value="<%=cemail%>" style="width: 261px" ><br><br>
	Subject: <br><input type="text" id="messagesubject" value="A copy of your invoice is attached" style="width: 261px" >
<br><br>
	Message to Customer<br>
<textarea style="width:300px;height:100px;" id="emailinvoicemessage">Your repair invoice from <%=request.cookies("shopname")%> is attached.  Thank you for your business!</textarea>
<br><br>
	<br><input id="updateemail" value="on" name="Checkbox1" type="checkbox"> 
	Update customer info with this email address<br>
	<br>&nbsp;<input onclick="redirectWithEmail()" name="Button4" type="button" value="Send Email"> <input onclick="document.getElementById('sendemail').style.display='none';document.getElementById('hider').style.display='none'" name="Button1" type="button" value="Cancel" ></div>

<div style="text-align:center;border:15px black solid;color:red;background-color:white;width:90%;height:90%;position:absolute;top:5%;left:5%;z-index:1050;display:none" id="invoicediv">
	<strong>Your Invoice is displayed below.</strong>
<input onclick="closePrint()" name="Button1" type="button" value="Close" > <input onclick="showEmailAddress('no')" name="Button1" type="button" value="E-mail" > <input onclick="printIframe('invoice')" id="printbutton" value="Print" type="button">
<iframe frameborder="0" scrolling="no" style="width:95%;height:95%;z-index:1599" id="invoice" name="invoice"></iframe>

</div>

<%
if esign = "yes" then
%>
<style type="text/css">
#printdrop{
	display:none;
	position:absolute;
	width:100%;
	background-color:#F0F0F0;
	z-index:1001;
	font-weight:bold; 
	left: 0px; 
	top: 50px; 
	padding:0px;
	border:3px #0066CC solid;
	
}
.printmenu{
	padding:5px;
	cursor:pointer;
	text-align:center;
	border-bottom:3px #0066CC solid;
	width:20%;
	border:3px #0066cc solid;
}
.printmenu:hover{
	color:black;
	background-color:white;
}
.auto-style1 {
	border-style: solid;
}
</style>

<table style="width:60%;margin-left:20%" id="printdrop" cellpadding="2" cellspacing="0" class="auto-style1">
	<tr>
		<td onclick="document.getElementById('printdrop').style.display='none';document.getElementById('hider').style.display='none'" class="printmenu"><img src="newimages/close-icon.png" height="30" width="27"><br/>
		Close&nbsp;</td>
	</tr>
	<tr>
		<td colspan="5">
		
		<div style="width:100%;" id="printedrolist">
			Any E-Signed RO's Are listed here:<br><br>
		<%
		'on error resume next
		set fso = server.createobject("scripting.filesystemobject")
		folpath = "d:\savedinvoices\" & shopid & "\" & request("roid")
		'folpath = server.mappath(folpath)
		'response.write folpath
		if fso.folderexists(folpath) then
			fol = fso.getfolder(folpath)
			if fso.folderexists(fol) then
				'response.write folpath
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
					if left(filename,2) = "ro" then
						far = split(filename,"_")
						' 2-7
						filedate = far(2) & "/" & far(3) & "/" & far(4) & " at " & far(5) & ":" & far(6) & ":" & far(7)
						'response.write "filename:" & filename
						response.write "<div onclick='reprintRO(" & chr(34) & f.name & chr(34) & ")' style='text-align:center;padding:4px;color:#3366CC;cursor:pointer;background-color:" & bgcolor & "'>Invoice printed and dated " & filedate & "</div><br>"
					else
						filedate = f.datelastmodified
						response.write "<div onclick='reprintRO(" & chr(34) & f.name & chr(34) & ")' style='text-align:center;padding:4px;color:#3366CC;cursor:pointer;background-color:" & bgcolor & "'>Invoice printed and dated " & filedate & "</div><br>"

					end if
				next
			else
				response.write "No previously signed RO's on file. <a href='viewro.asp?roid=" & request("roid") & "&printro=yes'>Click Here to Print an Unsaved Repair Order</a>"
			end if
		else
			response.write "No previously signed RO's on file. <a href='viewro.asp?roid=" & request("roid") & "&printro=yes'>Click Here to Print an Unsaved Repair Order</a>"
		end if
		%>
		</div>
		</td>
	</tr>
</table>
<%end if%>
<iframe id="invarchframe" frameborder="0" src='invoicearchives/showinvoices.asp?roid=<%=request("roid")%>' style="border:medium black outset;width:90%;height:90%;position:absolute;left:5%;top:5%;display:none;z-index:999"></iframe>

<table border=2 cellpadding=0 cellspacing=0 width=100% bordercolor="#0066CC" style="border-collapse: collapse"><tr><td>
<form name="theform" method="post" action="savero.asp">
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
            <td height="22" ><table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
               <td width="100%" align="center" class="style10" ><b><i>
               <font color="#FFFF00" size="5">RO#&nbsp;<%=roid%></font></i></b><img style="position:absolute;right:5px;top:0px;" alt="" height="37" src="newimages/shopbosslogo-ro.png" width="157"></td>
              </tr>
              <tr>
               <td width="100%" align="left" bgcolor="#0066CC">

             
               <table style="width: 100%">
				   <tr>
				   	   <%
				   	   if esign = "no" then
				   	   %>
					   <td onclick="location.href='pdfinvoices/printpdfro.asp?roid=<%=request("roid")%>'" valign="middle" class="menustyle">
					   <img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br>
					   Print RO</td>
					   <td onclick="document.getElementById('invarchframe').style.display='block';document.getElementById('hider').style.display='block'" valign="middle" class="menustyle"><img height="30" src="newimages/invoice-icon.png" width="30" >
					   <br>Archived Invoices</td>
					   <%
					   else
					   %>
					   <td onclick="showSavedROs()" valign="middle" class="menustyle">
					   <img alt="" height="30" src="newimages/printer_icon.gif" width="33"><br>
					   Print RO</td>
					   <%
					   end if
					   %>
					   <td onclick="showHistory()" class="menustyle">
					   <img alt="" height="30" src="newimages/historyx30.png" width="30"><br>
					   History</td>
					   <td onclick="location.href='customeredit.asp?customerid=<%=customerid%>'" class="menustyle">
					   <img src="http://<%=request.servervariables("SERVER_NAME")%>/sbp/newimages/peopleicon.png"><br>
					   Customer</td>
					   <td onclick="showReopenHistory()" class="menustyle">
					   <img height="30" src="newimages/open-file.png" width="30"><br>
					   RO Re-Open History</td>
					   <td onclick="showPics()" class="menustyle">
					   <img alt="" height="30" src="newimages/camera.jpg" width="35"><br>
					   View Pics</td>
					   <td onclick="<%=newsp%>" class="menustyle">
					   <img alt="" height="30" src="newimages/save.png" width="30">&nbsp;<br>
					   Exit</td>
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
              <font size="2" face="Arial">Customer:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><font size="2" face="arial"><%=rs("Customer")%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" class="style14">
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
              <td style="font-family:Arial;font-size:x-small;font-weight:normal" width="10%" align="right" valign="bottom" class="style1">
              Bus. Contact</td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" bordercolor="#000000" class="style15"><%=rs("contact")%>
			  &nbsp;</td>
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
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">License:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("VehLicNum")%>
			  &nbsp;&nbsp;&nbsp;Fleet #<%=rs("fleetno")%></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">VIN:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("Vin")%></font>
			  &nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Miles In:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15">
               <font size="2" face="arial">
               <%=rs("VehicleMiles")%></b></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Miles Out:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15">
               <%=rs("MilesOut")%>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Engine/Cyl:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("VehEngine")%>
			  &nbsp;&nbsp;<%=rs("Cyl")%>Cyl</font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" class="style14">
              <font size="2" face="Arial">Trans/Drive:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="top" class="style15"><font size="2" face="arial"><%=rs("VehTrans")%>
			  &nbsp;&nbsp;&nbsp;&nbsp;<%=rs("DriveType")%></font></td>
             </tr>
            </table></td>
           <td width="33%" valign="top" height="170">
           <table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" height="100%">
             <tr>
              <td width="10%" align="right" valign="top" height="44" style="height: 22px" class="style14">
			  RO Date</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="44" style="height: 22px" class="style15">
			  Open: <%=rs("datein")%> &nbsp;Closed:<%=rs("statusdate")%></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" height="44" style="height: 22px" class="style14">
			  Status</td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="44" style="height: 22px" class="style15">
			  <%
              	if isnumeric(left(rs("status"),1)) then
              		displays = right(s,len(s)-1)
              	else
              		displays = s
              	end if
				response.write s
			  %>
			  </td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" height="31" class="style14"><font size="2" face="arial">
			  Comments:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="31" class="style15">
			  <%=rs("Comments")%>&nbsp;</td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="top" height="22" class="style14"><font size="2" face="arial">
			  Writer:</font></td>
              <td width="0%" bgcolor="#FFFFFF" valign="top" height="22" class="style15">
               <%=rs("Writer")%>&nbsp;</td>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14"><font size="2" face="arial">
			  Source:</font></td>
              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15"><font size="2" face="arial">
               <%=rs("Source")%></font></td>
             </tr>
             <tr>
              <td width="10%" align="right" valign="bottom" height="22" class="style14">
               <p align="right" class="style13"><font size="2" face="arial">
			   Type:</font></td>


              <td width="20%" bgcolor="#FFFFFF" valign="bottom" height="22" class="style15">
			  <%=rs("ROType")%>&nbsp;</td>
             </tr>
            </table></td>
          </tr>
          <tr>
           <td style="max-height:500px;overflow-y:scroll" width="67%" valign="top" bgcolor="#FFFFFF" colspan="2">

           <!-- #include file=complaintlayoutview.asp -->

           </td>
           <td width="33%" valign="top" >
           <table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse; height:30px;" bordercolor="#111111">
             <tr>
              <td style="cursor:pointer" id="revcell" Width="20%" align="center" onclick="revChg()" class="style9">
			  <img alt="" height="30" src="newimages/desktop.png" width="30"><br>
			  Revisions</td>
              <td style="cursor:pointer" id="feecell" width="20%" align="center" onclick="feesChg()" class="style9">
			  <img alt="" height="30" src="newimages/discount_icon.png" width="28"><br>
			  Fees/Disc</td>
              <td style="cursor:pointer" id="warrcell" width="20%" align="center" onclick="warrChg()" class="style9">
			  <img alt="" height="30" src="newimages/warricon2.png" width="32"><br>
			  Warranty</td>
              <td style="cursor:pointer" id="pmtscell" width="20%" align="center"onclick="pmtsChg()" class="style9">
			  <img alt="" height="30" src="newimages/payment.png" width="30"><br>
			  Payments</td>
              <td style="cursor:pointer; width: 20%;" id="commcell" width="20%" align="center"onclick="changeCells('comm')" class="style9">
			  <img alt="" height="30" src="newimages/commlog.png" width="30"><br>
			  Comm Log</td>

             </tr>
            </table><table cellspacing="0" cellpadding="0" width="100%" border="0">
             <tr>
              <td id="est" style="display:block">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td width="50%" colspan="2" class="style17">
                  <p align="right" class="style13">Original Estimate:</p>
                 </td>
                 <td width="50%" colspan="2" bgcolor="#FFFFFF" class="style15"><%=formatnumber(rs("OrigRO"),2)%>
				 &nbsp;</td>
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
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=formatnumber(rs("Rev1Amt"),2)%>
				 &nbsp;</td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Amt:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=formatnumber(rs("Rev2Amt"),2)%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev1Date")%>
				 &nbsp;</td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Date:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev2Date")%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev1Phone")%>
				 &nbsp;</td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Phone:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev2Phone")%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev1Time")%>
				 &nbsp;</td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 Time:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev2Time")%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="25%" align="right" class="style17"><font size="2">
				 By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev1By")%>
				 &nbsp;</td>
                 <td width="25%" align="right" class="style17"><font size="2">
				 By:&nbsp;</font></td>
                 <td width="25%" bgcolor="#FFFFFF" class="style20"><%=rs("Rev2By")%>
				 &nbsp;</td>
                </tr>
               </table></td>
             </tr>
             <tr>
              <td id="fees" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">
				 Haz. Waste Fee:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
                  <p align="right"><%=formatdollar(rs("HazardousWaste"))%>&nbsp;</p>
                 </td>
                </tr>
  <center>
         <center>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2"><%=rs("userfee1label")%>
				 :&nbsp;</font></td>
                </center></center>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
                  <p align="right"><%=formatdollar(rs("UserFee1"))%>&nbsp;</p>
         </td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2"><%=rs("userfee2label")%>
				 :&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
				 <%=formatdollar(rs("UserFee2"))%>&nbsp;</td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">
				 Parts Tax Rate:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=rs("TaxRate")%><b><font size="3">
				 %</font></b></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14">Labor Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=rs("LaborTaxRate")%>
                 <b><font size="3">%</font></b></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14">Sublet Tax Rate:&nbsp; </td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15"><%=rs("SubletTaxRate")%>
                 <b><font size="3">%</font></b></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style14"><font size="2">
				 Discount Amt:&nbsp;</font></td>
                 <td width="50%" align="right" bgcolor="#FFFFFF" class="style15">
				 <%=formatdollar(rs("DiscountAmt"))%>&nbsp;</td>
                </tr>
                </table></td>
             </tr>
             <tr>
              <td id="warr" style="display:none">
              <table width="100%" cellspacing="0" cellpadding="2" bordercolor="#111111" class="style16">
                <tr>
                 <td width="50%" align="right" class="style17"><font size="2">
				 Warranty Months:&nbsp;</font></td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><%=rs("WarrMos")%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17"><font size="2">
				 Warranty Miles:&nbsp;</font></td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><%=rs("WarrMiles")%>
				 &nbsp;</td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
                 <td width="50%" bgcolor="#FFFFFF" class="style18"><input style="visibility:hidden; height:18" type="text" name="T26" size="7"></td>
                </tr>
                <tr>
                 <td width="50%" align="right" class="style17">&nbsp;</td>
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
             	<td style="display:none" id="comm">
              <table cellpadding="2" cellspacing="0" style="width: 100%">
				  <tr>
					  <td colspan="3" class="auto-style2" style="height: 20px;height:40px;background-color:maroon;text-align:center;color:white">
					  <strong>RO Communication Log</strong></td>
				  </tr>

				  <tr>
					  <td class="style32">Date/Time</td>
					  <td class="style32">By</td>
					  <td class="style32">Communication</td>
				  </tr>
				  <%
				  commlogstmt = "select * from repairordercommhistory where shopid = '" & shopid & "' and roid = " & request("roid")
				  set logrs = con.execute(commlogstmt)
				  if not logrs.eof then
				  	do until logrs.eof
				  %>
				  <tr>
					  <td><%=logrs("datetime")%>&nbsp;</td>
					  <td><%=logrs("by")%>&nbsp;</td>
					  <td><%=logrs("comm")%>&nbsp;</td>
				  </tr>
				  <%
				  		logrs.movenext
				  	loop
				  end if
				  %>

			  </table>
             	</td>
             </tr>
             <tr>
              <td id="pmts" style="display:none" class="style15">
              <table width="100%" border="1" cellspacing="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
                 <td width="50%" align="right" class="style22" style="width: 100%">
				 <table cellpadding="2" cellspacing="0" style="width: 100%">
					 <tr>
						 <td class="style24"><strong>Date</strong></td>
						 <td class="style24"><strong>Method</strong></td>
						 <td class="style23"><strong>Amount</strong></td>
					 </tr>
					 <%
					 tpmts = 0
					 stmt = "select * from accountpayments where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
					 set trs = con.execute(stmt)
					 if not trs.eof then
					 	do until trs.eof
					 		tpmts = tpmts + trs("amt")
					 %>
					 <tr>
						 <td><%=trs("pdate")%>&nbsp;</td>
						 <td><%=trs("ptype")%>&nbsp;</td>
						 <td class="style18"><%=formatcurrency(trs("amt"))%>&nbsp;</td>
					 </tr>
					 <%
					 		trs.movenext
					 	loop
					 else
					 %>
					 <tr>
						 <td colspan="3">No Payments Found&nbsp;</td>
					 </tr>
					<%
					end if
					%>
				 </table>
					</td>
                 </table></td>
             </tr>
            </table>
           <table width="100%" cellspacing="0" cellpadding="3" class="style16">
             <tr>
              <td style="padding:5px;" width="100%" colspan="2" class="style8">
               <strong>** Totals **</strong></td>
             </tr>

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
              <td width="50%" align="right" class="style21">Total Fees</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(rs("TotalFees"))%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Subtotal</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub)%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Tax</td>
              <%
              if rs("discountamt") > 0 then
              	if rs("discounttaxable") = "yes" and rs("taxrate") > 0 then
              		newttltax = formatnumber(rs("salestax") - (rs("discountamt") * (rs("taxrate")/100)),2)
              	else
              		newttltax = formatnumber(rs("salestax"),2)
              	end if
              else
              	newttltax = formatnumber(rs("salestax"),2)
              end if
              		
              %>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=newttltax%></b>
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
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+newttltax-cdbl(rs("DiscountAmt")))%></b>
			  &nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Payments</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15" style="color:red">
			  &lt; <%=formatnumber(tpmts,2)%> &gt;&nbsp;&nbsp;</td>
             </tr>
             <tr>
              <td width="50%" align="right" class="style21">Balance</td>
              <td width="50%" bgcolor="#FFFFFF" align="right" class="style15"><%=formatdollar(cdbl(rs("TotalFees"))+ttlprts+ttllbr+ttlsub+newttltax-cdbl(rs("DiscountAmt"))-tpmts)%>
			  &nbsp;&nbsp;</td>
             </tr>
             <center><center>
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
		response.write "<input type=hidden name='" & nrs.fields(fnum).name
	else
		response.write "<input type=hidden name='db" & nrs.fields(fnum).name
	end if
	if nrs.fields(fnum).name = "TotalLbr" then
		response.write "' value='" & ttllbr & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalPrts" then
		response.write "' value='" & ttlprts & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalSublet" then
		response.write "' value='" & ttlsub & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "SalesTax" then
		response.write "' value='" & nttltax & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "TotalRO" then
		response.write "' value='" & ttllbr+ttlprts+ttlsub+nttltax & "'>" & chr(10)
	elseif nrs.fields(fnum).name = "Subtotal" then
		response.write "' value='" & ttllbr+ttlprts+ttlsub & "'>" & chr(10)
	else
		response.write "' value='" & nrs(fnum) & "'>" & chr(10)
	end if
	response.write "<input type=hidden name='fldtype" & nrs.fields(fnum).name
	response.write "' value='" & nrs.fields(fnum).type & "'>" & chr(10)
next

'calc gp
ttls = ttlprts+ttllbr
ttlc = pc+tlc
dprofit = ttls-ttlc
if dprofit > 0 and ttls > 0 then gp = dprofit/ttls
if gp > 0 then gp = formatpercent(gp)
  %>
</form>
</td></tr></table>

<iframe frameborder="0" src="upload/upload.asp?roid=<%=request("roid")%>" style="border:medium black outset;width:60%;height:80%;position:absolute;left:20%;top:10%;display:none;z-index:999" id="showframe"></iframe>
<iframe id="picframe" frameborder="0" src='upload/showpics.asp?s=c&roid=<%=request("roid")%>' style="border:medium black outset;width:90%;height:90%;position:absolute;left:5%;top:5%;display:none;z-index:999"></iframe>




<font color="#FFFFFF">


<div id="hider"></div>
<div id="timer">
<img src="newimages/ajax-loader.gif">
</div>

<div style="text-align:center;border:15px black solid;color:red;background-color:white;width:90%;height:90%;position:absolute;top:5%;left:5%;z-index:1050;display:none" id="invoicediv">
	<strong>Your Invoice is displayed below.&nbsp; Move your mouse to the bottom to 
	print it.&nbsp; Click here to</strong>
<input onclick="closePrint()" name="Button1" type="button" value="Close" >

<iframe frameborder="0" scrolling="no" style="width:95%;height:99%;" id="invoice" name="invoice"></iframe>
</div>

<script language="javascript">
	status = "Shop Boss Pro"
<%
if request("printro") = "yes" then
%>

function closePrint(){

	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("hider").style.display = "none"

}

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
	var url="pdfinvoices/printpdfro.asp?roid=<%=request("roid")%>"
	console.log(url)
	xmlHttp.onreadystatechange=printROStateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function printROStateChanged(){
	document.getElementById("hider").style.display = "block"
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){

		//alert(xmlHttp.responseText)
		invname = xmlHttp.responseText
		document.getElementById("hider").style.display = "block"
		document.getElementById("invoicediv").style.display = "block"
		document.getElementById("invoice").src = "pdfinvoices"+invname
		document.getElementById('timer').style.display = 'none'
		
	}
}

var xmlHttp

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

<%
end if
%>

</script>


<iframe src="rohistorydata.asp?roid=<%=request("roid")%>&vid=<%=myvid%>" style="display:none" id="history"></iframe>
<script language="javascript">
	status = "Shop Boss Pro"
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>