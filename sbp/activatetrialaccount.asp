<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/connwoshopid.asp -->
<%


stmt = "select shopid from deniedlist where shopid = '" & request("shopid") & "'"
set rs = con.execute(stmt)
if not rs.eof then
	response.redirect "renewbyphone.asp"
end if
set rs = nothing

if cint(request("shopid")) >= 4277 then 
	response.redirect "activatetrialaccount-new.asp?shopid=" & request("shopid")
end if

Function PCase(ByVal strInput)' As String
    Dim I 'As Integer
    Dim CurrentChar, PrevChar 'As Char
    Dim strOutput 'As String

    PrevChar = ""
    strOutput = ""

    For I = 1 To Len(strInput)
        CurrentChar = Mid(strInput, I, 1)

        Select Case PrevChar
            Case "", " ", ".", "-", ",", """", "'"
                strOutput = strOutput & UCase(CurrentChar)
            Case Else
                strOutput = strOutput & LCase(CurrentChar)
        End Select

        PrevChar = CurrentChar
    Next 'I

    PCase = strOutput
End Function 

function dconv(d)
	on error resume next
	if len(d) > 0 then
		dar = split(d,"/")
		if len(dar(2)) = 2 then
			myear = "20" & dar(2)
		else
			myear = dar(2)
		end if
		mmonth = dar(0)
		if len(mmonth) = 1 then mmonth = "0" & mmonth
		mday = dar(1)
		if len(mday) = 1 then mday = "0" & mday
		dconv = myear&"-"&mmonth&"-"&mday
	else
		dconv = ""
	end if
	
end function

function oconv(d)

	dar = split(d,"-")
	oconv = dar(1) & "/" & dar(2) & "/" & dar(0)

end function

function newOConv(d)

	dar = split(d,"-")
	if len(dar(1)) = 1 then 
		mmonth = "0" & dar(1)
	else
		mmonth = dar(1)
	end if
	if len(dar(2)) = 1 then
		dday = "0" & dar(2)
	else
		dday = dar(2)
	end if
	yyear = dar(0)
	newOConv = mmonth & "/" & dday & "/" & yyear
	

end function

function formatPhone(p)
	x = fixPhone(p)
	if len(x) = 10 then
		formatphone = "(" & left(x,3) & ") " & mid(x,4,3) & "-" & right(x,4)
	elseif len(x) = 7 then
		formatphone = left(x,3) & "-" & right(x,4)
	else
		formatphone = ""
	end if

end function

function fixPhone(p)

	x = replace(p,")","")
	x = replace(x,"(","")
	x = replace(x,"-","")
	x = replace(x," ","")
	x = replace(x,".","")
	fixPhone = x
	
end function

if len(request("cc")) > 0 then

	stmt = "update company set blastname = '" & request("blastname") & "', bfirstname = '" & request("bfirstname") & "', howpaying = 'cc'" _
	& ", cc = to_base64(AES_ENCRYPT('" & request("cc") & "','kdvk7k2aG3TKidQmCu4kLmtgRNaBXyR1!$HDN$@==**JDNdjdye7564')), expdate = '" & request("expdate") & "', billingaddress = '" _
	& request("billingaddress") & "', billingstate = '" & request("billingstate") & "', billingcity = '" & request("billingcity") & "', billingzip = '" _
	& request("billingzip") & "', `status` = 'ACTIVE'" _
	& ", package = 'Paid', dateofacceptance = '" & dconv(date) & "', initialdoa = '" & dconv(date) & "' where shopid = '" & request("shopid") & "'"
	'response.write stmt
	con.execute(stmt)
		

	cmess = "<span style='color:green;font-weight:bold;font-size:16px;'>Your account information has been updated.  Thank you</span>"

	set mail = server.createobject("persits.mailsender")
	mail.host = "smtp.emailsrvr.com"
	mail.from = "support@shopbosspro.com"
	mail.fromname = "Activated Account"
	mail.subject = "Activated Account"
	mail.addaddress "support@shopbosspro.com"
	mail.username = "support@shopbosspro.com"
	mail.password  = "ashley92"
	mb = "The following Shop has activated their account.  You do not need to charge their account." & chr(10) & chr(10)
	
	'get shop info
	stmt = "select companyname from company where shopid = '" & request("shopid") & "'"
	set rs = con.execute(stmt)
	mb = mb & rs("companyname") & chr(10)
	mb = mb & request("shopid") & chr(10)
	
	mail.body = mb
	mail.send
	response.redirect "login.asp?shopnum=" & request("roid")
end if

'get shop name
stmt = "select * from company where shopid = '" & request("shopid") & "'"
set rs = con.execute(stmt)
sn = rs("companyname")

if cint(request("shopid")) >= 4277 then
	response.redirect "activatetrialaccount-new.asp?shopid=" & request("shopid")
end if
%>
<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Untitled 1</title>
<style type="text/css">
.cmenu{
	background-color:#F0F0F0;
	text-align:center;
	color:#800000;
	font-size:12px;
	font-weight:bold;
	border:1px gray outset;
	cursor:pointer;
	height:80px;
	width:7.1%;
	font-family:Arial;
}

.style5 {
	font-family: Arial, Helvetica, sans-serif;
}

.style6 {
	text-align: center;
	color: #FFFFFF;
}

.style7 {
	text-align: center;
	color: #000000;
	font-size: small;
}

.style8 {
	color: #FF0000;
}
.style9 {
	font-family: Arial, Helvetica, sans-serif;
	text-align: center;
}

</style>
<%
function convDS(d)

	dar = split(d,"/")
	if len(dar(0)) = 1 then
		mmonth = "0" & dar(1)
	else
		mmonth = dar(0)
	end if
	if len(dar(1)) = 1 then
		dday = "0" & dar(1)
	else
		dday = dar(1)
	end if
	yyear = dar(2)
	
	d = yyear & "-" & mmonth & "-" & dday


end function

stmt = "select * from company where shopid = '" & request("shopid") & "'"
set rs = con.execute(stmt)
if rs.eof then
	response.redirect "login1.asp"
else
	shopname = rs("companyname")
	baddress = rs("billingaddress")
	bcity = rs("billingcity")
	bstate = rs("billingstate")
	bzip = rs("billingzip")
	susdate = rs("datesuspended")
	
	if lcase(rs("estguide")) = "yes" then
		if rs("alldatarepair") = "yes" then
			alldatachg = 99.95
		else
			alldatachg = 29.95
		end if
	else
		alldatachg = 0
	end if
	
	ds = convDS(rs("datestarted"))
	ds = datediff("d","6/03/2013",rs("datestarted"))

	'get the month owed
	
	if not isnull(susdate) then
		dateowed = dateadd("d",-30,susdate)
		monthdue = cint(datepart("m",dateowed))
		if monthdue = "12" then
			yeardue = cint(datepart("yyyy",susdate)) - 1
		else
			yeardue = datepart("yyyy",susdate)
		end if
	else
		'get last ro date and assume month prior
		stmt = "select roid,datein from repairorders where shopid = '" & request("shopid") & "' order by datein desc limit 1"
		'response.write stmt & "<BR>"
		set trs = con.execute(stmt)
		if not trs.eof then
			lro = trs("datein")
			
			dateowed = dateadd("d",-0,lro)
			'response.write "dateowed:" & dateowed & "<BR>"
			monthdue = cint(datepart("m",dateowed))
			yeardue = datepart("yyyy",dateowed)
			'response.write "lro:" & lro & "<BR>"
		else
			'response.write "renewbyphone.asp"
		end if
		
	end if
	
	select case monthdue
		case 1
			eom = "31"
		case 2
			eom = "28"
		case 3
			eom = "31"
		case 4
			eom = "30"
		case 5
			eom = "31"
		case 6
			eom = "30"
		case 7
			eom = "31"
		case 8
			eom = "31"
		case 9
			eom = "30"
		case 10
			eom = "31"
		case 11
			eom = "30"
		case 12
			eom = "31"
	end select
	
	if len(monthdue) = 1 then monthdue = "0" & cstr(monthdue)
	
	'get amount owed
	sd = yeardue & "-" & monthdue & "-01"
	ed = yeardue & "-" & monthdue & "-" & eom
	
	stmt = "select count(*) as c from repairorders where shopid = '" & request("shopid") & "' and datein >= '" & sd & "' and datein <= '" & ed & "'"
	'response.write stmt & "<BR>"
	set trs = con.execute(stmt)
	if not isnull(trs("c")) then
		rocount = cdbl(trs("c"))
	else
		rocount = 0
	end if

	if ds <= 0 then
		if rocount > 0 and rocount <= rs("freelimit") then
			chg = 0
		elseif rocount > rs("freelimit") and rocount <= 10 then
			chg = 14.95
		elseif rocount > rs("freelimit") and rocount <= 30 then
			chg = 29.95
		elseif rocount >= 31 and rocount <= 55 then
			chg = 69.95
		elseif rocount > 55 then
			chg = 99.95
		end if
	end if
	if ds > 0 then
		if rocount > 0 and rocount <= rs("freelimit") then
			chg = 0
		elseif rocount > rs("freelimit") and rocount <= 30 then
			chg = 29.95
		elseif rocount >= 31 and rocount <= 55 then
			chg = 69.95
		elseif rocount > 55 then
			chg = 99.95
		end if
	end if
	
	totalchg = chg + alldatachg
	
	'response.write rocount & ":" & totalchg
end if
%>

<script language="javascript">
function changeCharge(v){
	d0 = document.getElementById("cc0")
	d1 = document.getElementById("cc1")
	d2 = document.getElementById("cc2")
	d3 = document.getElementById("cc3")
	d4 = document.getElementById("cc4")
	d5 = document.getElementById("cc5")
	d6 = document.getElementById("cc6")
	b1 = document.getElementById("ba1")
	b2 = document.getElementById("ba2")
	b3 = document.getElementById("ba3")
	b4 = document.getElementById("ba4")
	if(v == "cc"){
		b1.style.visibility = "hidden"
		b2.style.visibility = "hidden"
		b3.style.visibility = "hidden"
		b4.style.visibility = "hidden"
		d0.style.visibility = "visible"
		d1.style.visibility = "visible"
		d2.style.visibility = "visible"
		d3.style.visibility = "visible"
		d4.style.visibility = "visible"
		d5.style.visibility = "visible"
		d6.style.visibility = "visible"
	}
	if(v == "mobank"){
		b1.style.visibility = "visible"
		b2.style.visibility = "visible"
		b3.style.visibility = "visible"
		b4.style.visibility = "visible"
		d0.style.visibility = "hidden"
		d1.style.visibility = "hidden"
		d2.style.visibility = "hidden"
		d3.style.visibility = "hidden"
		d4.style.visibility = "hidden"
		d5.style.visibility = "hidden"
		d6.style.visibility = "hidden"
	}

}

function checkForm(){


	if(document.bankform.cc.value == ""){
		alert("Credit Card Number is required")
		return
	}
	if(document.bankform.expdate.value == ""){
		alert("Credit Card Expiration Date is required")
		return
	}
	if(document.bankform.billingaddress.value == ""){
		alert("Billing Address is required")
		return
	}
	if(document.bankform.blastname.value == ""){
		alert("Billing Last Name is required")
		return
	}
	if(document.bankform.bfirstname.value == ""){
		alert("Billing First Name is required")
		return
	}
	if(document.bankform.billingcity.value == ""){
		alert("Billing City is required")
		return
	}
	if(document.bankform.billingstate.value == ""){
		alert("Billing State is required")
		return
	}
	if(document.bankform.billingzip.value == ""){
		alert("Billing Zip Code is required")
		return
	}
	bexp = document.bankform.expdate.value
	if (bexp.length != 5){
		alert("Please enter the Expiration date on card as MM/YY")
		return false;
	}
	document.bankform.submit()

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

function postCCPmt(fname,lname,zip,ctype,cnum,expmo,expyear,shopid,amt,baddress,bcity,bstate){
	
	xmlHttp=GetXmlHttpObject()

	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	
	qs = "?baddress="+baddress+"&bcity="+bcity+"&bstate="+bstate+"&updatebilling=y&amt="+amt+"&fname="+fname+"&lname="+lname+"-"+shopid+"&zip="+zip+"&ctype="+ctype+"&cnum="+cnum+"&expmo="+expmo+"&expyear="+expyear+"&shopid=<%=request("shopid")%>"
	r=Math.floor(Math.random()*111111)
	qs += "&r=" + r
	var url="control/cc/cc_purchase.asp"+qs
	//alert(url)
	xmlHttp.onreadystatechange=stateChanged 
	xmlHttp.open("GET",url,true)
	xmlHttp.send(null)
}

function stateChanged(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete"){
		x = xmlHttp.responseText
		//alert(x)
		if(x == "approved"){
			alert("Payment approved and posted")
			//location.href='billing.asp?sd=<%=request("sd")%>&ed=<%=request("ed")%>'
			location.href = 'login1.asp'
		}else{
			alert(x)
		}
	}
}

</script>

</head>
<body>

<div style="text-align:center;height: 58px;background-image:url('newimages/rosubheader.jpg');font-family:Arial;">
<br/><span class="style6"><strong>You have reached the end of your Trial 
	Account.&nbsp; To activate your account, please update your 
	credit card or bank account information below</strong></span><br/>

</div>

<p>
<table cellpadding="6" cellspacing="0" style="width: 100%">
	<tr>
		<td style="width: 50%; ">
<form action="activatetrialaccount.asp" name="bankform" method="post">
	<input type="hidden" name="totalchg" value="<%=totalchg%>">
<input name="dateofacceptance" value="<%=date%>" type="hidden" />
<input name="howpaying" value="mobank" type="hidden" />
<input name="shopid" value="<%=request("shopid")%>" type="hidden" />

	<table style="width: 100%">
		<tr>
			<td class="style9" colspan="2"><strong><%=ucase(sn)%></strong>&nbsp;<br/><span class="style8"><strong>
			Your 30 day trial account has concluded.&nbsp; If you would like to 
			activate your account and <br />
			continue to use Shop Boss, please complete the information 
			below.<br />
			<br />
			Be sure that information shown below is accurate.<br />
			The information entered will be retained for future billing.</strong></span></td>
		</tr>
		<tr style="<%=cchide%>" id="cc0">
			<td class="style5" style="width: 537px"><b>Visa / Mastercard / 
			Discover / American Express</b></td>
			<td class="style5"></td>
		</tr>
		<tr style="<%=cchide%>" id="cc1">
			<td class="style5" style="width: 537px">Credit Card Number 
			(XXXX-XXXX-XXXX-XXXX)</td>
			<td class="style5">
<input name="cc" value=''  type="text" class="style5" style="width: 175px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc2">
			<td class="style5" style="width: 537px">Expiration Date (MM/YY)</td>
			<td class="style5">
<input name="expdate" value=''  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc2">
			<td class="style5" style="width: 537px">First Name on Card</td>
			<td class="style5">
<input name="bfirstname"  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc2">
			<td class="style5" style="width: 537px">Last Name on Card</td>
			<td class="style5">
<input name="blastname"  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc2">
			<td class="style5" style="width: 537px">Shop Name</td>
			<td class="style5"><%=shopname%>
			&nbsp;</td>
		</tr>
		<tr style="<%=cchide%>" id="cc3">
			<td class="style5" style="width: 537px">Billing Address</td>
			<td class="style5">
<input name="billingaddress" value=''  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc4">
			<td class="style5" style="width: 537px">Billing City</td>
			<td class="style5">
<input name="billingcity" value=''  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc5">
			<td class="style5" style="width: 537px">Billing State</td>
			<td class="style5">
<input name="billingstate" value=''  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc6">
			<td class="style5" style="width: 537px">Billing Zip</td>
			<td class="style5">
<input name="billingzip" value=''  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		</table>
<br />
<input name="Submit1" type="button" onclick="checkForm()" value="Activate My Account" /> <input name="Button1" type="button" value="Cancel" onclick="location.href='login1.asp'" /></p>
		</form></td>
	</tr>
</table>

	</p>
	<p><br />
</p>

</body>

</html>
