<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/connwoshopid.asp -->
<%
shopid = base64decode(request("shopid"))
roid = base64decode(request("roid"))
response.cookies("shopid") = shopid

response.redirect "https://" & request.servervariables("server_name") & "/status.php?shopid=" & shopid & "&roid=" & roid

on error resume next
if request("rs") <> "y" then
	'response.redirect "../autherror.asp"
end if
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

'response.cookies("asessid") = ""

'stmt = "select companyname,companyemail,logo,invoicetitle,estimatetitle from company where shopid = '" & shopid & "'"
'set rs = con.execute(stmt)
'shopname = rs("companyname")
'etitle = rs("estimatetitle")
'ititle = rs("invoicetitle")
'companyemail = rs("companyemail")
'if len(rs("logo")) > 0 then
'	logofile = "upload/" & shopid & "/" & rs("logo")
'end if

%>

<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Untitled 1</title>
<script type="text/javascript">

function getRO(){
	
	document.getElementById("spinner").style.display = "block"
    var xmlhttp = new XMLHttpRequest();
    xmlhttp.onreadystatechange = function() {
        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
        	document.getElementById("spinner").style.display = "none"
        }
    }
    // check for custom ro
    <%
    set fso = server.createobject("scripting.filesystemobject")
    shopfolder = server.mappath(".") & "\pdfinvoices\" & shopid
    if fso.folderexists(shopfolder) then
    %>
    var url="https://<%=request.servervariables("SERVER_NAME")%>/sbp/pdfinvoices/<%=shopid%>/printpdfro.asp?remoteauth=yes&shopid=<%=shopid%>&roid=<%=roid%>"
    <%
    else
    %>
    var url="https://<%=request.servervariables("SERVER_NAME")%>/sbp/pdfinvoices/printpdfro.asp?remoteauth=yes&shopid=<%=shopid%>&roid=<%=roid%>"
    <%
    end if
    %>
    //console.log(url)
    xmlhttp.open("GET", url, true);
    xmlhttp.send();

}


function signRO(){
	
	rtc = document.getElementById("retrycount").value
	console.log("rtc:"+rtc)
	if (rtc >= 2){
		alert("We're sorry but you have exceeded the limit of attempts to authenticate.")
		location.href = 'https://<%=request.servervariables("SERVER_NAME")%>/autherror.asp'
		return;
	}
	window.scroll(0,0)
	var a
	if (document.getElementById("address1").checked){
		a = document.getElementById("address1").value
	}
	if (document.getElementById("address2").checked){
		a = document.getElementById("address2").value
	}
	if (document.getElementById("address3").checked){
		a = document.getElementById("address3").value
	}
	if (document.getElementById("address4").checked){
		a = document.getElementById("address4").value
	}
	if (document.getElementById("address5").checked){
		a = document.getElementById("address5").value
	}
	if (document.getElementById("address6").checked){
		a = document.getElementById("address6").value
	}
	if (document.getElementById("address7").checked){
		a = document.getElementById("address7").value
	}
	if (document.getElementById("address8").checked){
		a = document.getElementById("address8").value
	}
	if (document.getElementById("address9").checked){
		a = document.getElementById("address9").value
	}
	if (document.getElementById("address10").checked){
		a = document.getElementById("address10").value
	}
	a = encodeURI(a)
	roid = document.getElementById("roid").value
	shopid = document.getElementById("shopid").value
	last4phone = document.getElementById("ph").value
	sig = document.getElementById("sig").value
	agreecheck = document.getElementById("agreecheck").checked
	
	if(agreecheck){
		document.getElementById("spinner").style.display = "block"
		data="roid="+roid+"&shopid="+shopid+"&ph="+last4phone+"&sig="+sig+"&a="+a
	    var xmlhttp = new XMLHttpRequest();
	    xmlhttp.onreadystatechange = function() {
	        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
	        	results = xmlhttp.responseText
	        	//alert("ropath:"+results)
	        	if (results == "programerror"){
	        		document.getElementById("spinner").style.display = "none"
	        		alert("There was an error generating the invoice")
	        		return
	        	}
	        	if (results == "validationerror"){
	        		if (rtc == 0){
		        		alert("You selected an incorrect address or entered an incorrect phone number.  You have two more tries")
		        	}
		        	if (rtc == 1){
		        		alert("You selected an incorrect address or entered an incorrect phone number.  You have one more try")
		        	}

	        		document.getElementById("spinner").style.display = "none"
	        		nrtc = parseFloat(rtc) + 1
	        		document.getElementById("retrycount").value = nrtc
	        		return
	        	}else{
	        		setTimeout(function(){
	        		//document.getElementById("invoice").src = "https://<%=request.servervariables("SERVER_NAME")%>/sbp/"+results
	        		//document.getElementById("invoice").style.width = "100%"
	        		document.getElementById("spinner").style.display = "none"
	        		document.getElementById("invresults").value = document.getElementById("invresults").value + results
	        		window.scroll(0,0)
	        		// hide the info box and show the success box.
	        		document.getElementById("info").style.display = "none"
	        		document.getElementById("header").style.display = "none"
	        		document.getElementById("unsigneddoc").style.display = "none"
	        		document.getElementById("success").style.display = "block"
	        		},2000)
	        	}
	        }
	    }
	    url = "auth2mobile.asp?"+data
	    //alert(url)
	    xmlhttp.open("GET", url, true);
	    xmlhttp.send();
	}
}


function validateEmail(email) {
    var re = /^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
    return re.test(email);
}

function sendEmail(){

	e = document.getElementById("customeremail").value
	if (validateEmail(e)){
		// send a copy to the customer via authemail.asp
	    var xmlhttp = new XMLHttpRequest();
	    xmlhttp.onreadystatechange = function() {
	        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
	        	results = xmlhttp.responseText
	        	//alert(results)
	        	if (results == "error"){
	        		alert("There was an error emailing the repair order")
	        	}else{
	        		if (results == "success"){
		        		alert("A copy of the email has been sent.  Thank you")
		        		location.href = 'https://www.google.com'
		        	}else{
		        		console.log(results)
		        	}
	        	}
	        }
	    }
	    iloc = document.getElementById("invresults").value.replace("https://<%=request.servervariables("SERVER_NAME")%>/sbp","")
	    
	    data = "mobile=y&customeremail="+e+"&shopemail=<%=companyemail%>&invlocation="+iloc+"&shopname=<%=server.urlencode(shopname)%>"
	    url = "authemail.asp?"+data
	    //alert(url)
	    xmlhttp.open("GET", url, true);
	    xmlhttp.send();

	}else{
		alert("You have entered and invalid email address")
	}

}
</script>
<style type="text/css">
body{
	font-family:Arial, Helvetica, sans-serif
}
.auto-style1 {
	font-size: 36pt;
}
.auto-style2 {
	text-align: center;
}
.auto-style3 {
	font-size: x-large;
}
.auto-style4 {
	text-align: right;
}
</style>
</head>
<%
function randomNumber()
	Dim max,min
	max=500
	min=1
	Randomize
	randomNumber = Int((max-min+1)*Rnd+min)
end function
num1 = randomNumber
num2 = randomNumber
num3 = randomNumber
num4 = randomNumber
num5 = randomNumber
num6 = randomNumber
num7 = randomNumber
num8 = randomNumber
num9 = randomNumber
num10 = randomNumber

function tempformatPhone(p)
	x = fixPhone(p)
	if len(x) = 10 then
		tempformatphone = "(" & left(x,3) & ") " & mid(x,4,3) & "-" & right(x,4)
	else
		tempformatPhone = x
	end if
end function

%>
<body onload="getRO()">
<input id="retrycount" type="hidden" value="0" />
<img id="spinner" style="display:none;position:absolute;left:0;right:0;bottom:0;margin:auto;top:100px;width:256px;height:256px;" alt="" src="newimages/loaderbig.gif" />
<%
stmt = "select * from randomstreet where id = " & num1
set rs = con.execute(stmt)
val1 = rs("street")
stmt = "select * from randomstreet where id = " & num2
set rs = con.execute(stmt)
val2 = rs("street")
stmt = "select * from randomstreet where id = " & num3
set rs = con.execute(stmt)
val3 = rs("street")
stmt = "select * from randomstreet where id = " & num4
set rs = con.execute(stmt)
val4 = rs("street")
stmt = "select * from randomstreet where id = " & num5
set rs = con.execute(stmt)
val5 = rs("street")
stmt = "select * from randomstreet where id = " & num6
set rs = con.execute(stmt)
val6 = rs("street")
stmt = "select * from randomstreet where id = " & num7
set rs = con.execute(stmt)
val7 = rs("street")
stmt = "select * from randomstreet where id = " & num8
set rs = con.execute(stmt)
val8 = rs("street")
stmt = "select * from randomstreet where id = " & num9
set rs = con.execute(stmt)
val9 = rs("street")

stmt = "select customeraddress,email,`status` from repairorders where shopid = '" & shopid & "' and roid = " & roid
set rs = con.execute(stmt)
if lcase(rs("status")) = "final" or lcase(rs("status")) = "closed" then
	etitle = ititle
else
	etitle = etitle
end if

val10 = rs("customeraddress")

set trs = server.createobject("adodb.recordset")
trs.fields.append "num", 3, 10
trs.fields.append "val", 200, 200
trs.open
trs.addnew array("num","val"), array(num1,val1)
trs.addnew array("num","val"), array(num2,val2)
trs.addnew array("num","val"), array(num3,val3)
trs.addnew array("num","val"), array(num4,val4)
trs.addnew array("num","val"), array(num5,val5)
trs.addnew array("num","val"), array(num6,val6)
trs.addnew array("num","val"), array(num7,val7)
trs.addnew array("num","val"), array(num8,val8)
trs.addnew array("num","val"), array(num9,val9)
trs.addnew array("num","val"), array(num10,val10)
trs.update
trs.sort="num"

%>
<div class="auto-style4">
	
	<table style="width: 100%" cellpadding="3" cellspacing="0">
		<tr>
			<td valign="top" rowspan="2">
			<img style="float: left" src="<%=logofile%>" align="top" />
			</td>
			<td valign="top"><br/><span style="padding:15px;border:2px navy solid;border-radius:10px">
			<img alt="" height="29" src="newimages/padlock2.png" width="22" align="top" />Secure 
			Signature by
	<img alt="" height="28" src="newimages/shopbosslogo-sm.png" width="139" align="top" /></span></td>
		</tr>
		<tr>
			<td valign="top">
			<input id="unsigneddoc" onclick="getRO()" name="Button5" type="button" value="Download <%=etitle%>" style="width: 502px; height: 83px; font-size:24pt" /></td>
		</tr>
	</table>
	
	</div>
<div id="header" style="text-align:center;font-size:18pt" class="auto-style1"><br/><strong><%=shopname%> is requesting an E-Signature on an <%=etitle%>.&nbsp; To 
add your electronic signature, please complete the information below.</strong></div>
<div style="display:none;font-size:36pt" id="success" class="auto-style2">

	<strong>Thank you for using our E-Signature system.<br />
	<input onclick="location.href=document.getElementById('invresults').value" id="signeddoc" name="Button5" type="button" value="Download Signed <%=etitle%>" style="width: 502px; height: 83px; font-size:24pt" /><br/>
	<br/>
If you would like a copy emailed to you, verify the correct email address belowclick the Send Email button below.

	<br />
	<br />
	Email Address: </strong>
	<input type="hidden" id="invresults" value="https://<%=request.servervariables("SERVER_NAME")%>/sbp/" />
	<br />
	<input name="Text1" id="customeremail" value="<%=rs("email")%>" style="width: 750px; height: 74px;font-size:36pt" type="text" /><br />
	<br />
	<input name="Button3" type="button" onclick="sendEmail()" value="Send Email" style="width: 395px; height: 95px;font-size:36pt" />&nbsp; <strong>
	<span style="font-size:36pt" class="auto-style3">or</span></strong>&nbsp;
	<input name="Button4" onclick="location.href='https://www.google.com'" style="width: 395px; height: 95px;font-size:36pt" type="button" value="Done" /></div>

<%
' generate the RO

%><br/><iframe id="invoice" style="width:0%;height:0px;display:none" src=""></iframe>
<div style="padding:10px;border:2px navy solid;border-radius:10px;min-height:540px;" id="info">
<p>To authenticate you as authorized to sign this document, please 
complete the following information.<strong><br />
<br />
</strong>Please select the street address on file for this repair:</p>


<input id="shopid" name="shopid" type="hidden" value="<%=request("shopid")%>" />
<input id="roid" name="roid" type="hidden" value="<%=request("roid")%>"  />

<%
trs.movefirst
c = 1
do until trs.eof
%>
<input style="width:70px;height:70px;" value="<%=ucase(trs("val"))%>" id="address<%=c%>" name="address" type="radio" /> &nbsp;<span style="font-size:24pt"><%=ucase(trs("val"))%></span><br />
<%
	c = c + 1
	trs.movenext
loop

%>
<br />

	<span class="auto-style1">Please enter the last 4 digits of any phone number associated with this repair:</span><br />
	<input name="ph" id="ph" type="text" style="width: 800px; height: 105px; font-size:48pt" /><br />
	<br />
	<span class="auto-style1">Type your name in the box below as your signature</span><br />
	<input name="sig" id="sig" style="font-size:36pt;width: 800px; font-family:'Trebuchet MS’, Helvetica, sans-serif'; height: 105px;" type="text" /><br />
	<br />
	<input id="agreecheck" name="Checkbox1" type="checkbox" style="width: 70px; height: 70px" /> 
	<span class="auto-style1">I hereby add my signature to this 
	<%=etitle%> and approve the repair work.&nbsp; I understand that this is a binding agreement 
	between myself and <%=ucase(shopname)%> </span> <strong>.</strong><br />
	<br />
&nbsp;<input name="Button1" type="button" onclick="signRO()" value="I Agree and Authorize" style="width: 876px; height: 113px; font-size:36pt" /><br />
	<br />
	<br />
&nbsp;<input name="Button2" type="button" value="Cancel - I DO NOT Authorize" style="width: 875px; height: 113px; font-size:36pt" />

</div>
</body>
<%
'response.cookies("asessid") = shopid
if err.number <> 0 then response.write err.description
%>
</html>
