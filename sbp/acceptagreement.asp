<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/connwoshopid.asp -->
<%

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

if len(request("paymethod")) > 0 then
	if request("paymethod") = "mobank" then
		stmt = "update company set howpaying = 'mobank', bankaccount = '" & request("bankaccount") & "', " _
		& "routing = '" & request("routing") & "', " _
		& "dateofacceptance = '" & dconv(request("dateofacceptance")) & "', " _
		& "last4 = '" & request("last4") & "', " _
		& "birthdate = '" & request("birthdate") _
		& "', estguide = 'No', " _
		& "package = 'Paid'" _
		& " where shopid = '" & replace(request.cookies("shopid"),"temp-","") & "'"
		response.write stmt
		con.execute(stmt)
		response.cookies("paid") = "yes"
		response.redirect "wip.asp"
		
		cmess = "<span style='color:green;font-weight:bold;font-size:16px;'>Your account information has been updated.  Thank you</span>"
	else
		stmt = "update company set dateofacceptance = '" & dconv(request("dateofacceptance")) & "', howpaying = '" & request("paymethod") & "'" _
		& ", cc = to_base64(AES_ENCRYPT('" & request("cc") & "','kdvk7k2aG3TKidQmCu4kLmtgRNaBXyR1!$HDN$@==**JDNdjdye7564')), expdate = '" & request("expdate") & "', billingaddress = '" _
		& request("billingaddress") & "', billingstate = '" & request("billingstate") & "', billingcity = '" & request("billingcity") & "', billingzip = '" _
		& request("billingzip") & "', estguide = 'No', package = 'Paid'" _
		& " where shopid = '" & replace(request.cookies("shopid"),"temp-","") & "'"
		'response.write stmt
		con.execute(stmt)
		response.cookies("paid") = "yes"
		response.redirect "wip.asp"

		cmess = "<span style='color:green;font-weight:bold;font-size:16px;'>Your account information has been updated.  Thank you</span>"
	end if
	set mail = server.createobject("persits.mailsender")
	mail.host = "smtp.emailsrvr.com"
	mail.from = "support@shopbosspro.com"
	mail.fromname = "New Account Agreement"
	mail.subject = "New Account Agreement"
	mail.addaddress "support@shopbosspro.com"
	mail.username = "support@shopbosspro.com"
	mail.password  = "ashley92"
	mb = "The following Shop has agreed to use Shop Boss:" & chr(10) & chr(10)
	
	'get shop info
	stmt = "select * from company where shopid = '" & replace(request.cookies("shopid"),"temp-","") & "'"
	set rs = con.execute(stmt)
	mb = mb & rs("companyname") & chr(10)
	mb = mb & rs("shopid") & chr(10)
	
	if request("lg") = "Yes" then
		mb = mb & "Include Labor Guide"
	elseif request("lg") = "No" then
		mb = mb & "DO NOT include Labor Guide"
	end if
	mail.body = mb
	mail.send
end if

if cint(replace(request.cookies("shopid"),"temp-","")) >= 4277 then 
	response.redirect "activatetrialaccount-new.asp?shopid=" & replace(request.cookies("shopid"),"temp-","")
end if

'get shop name
stmt = "select companyname from company where shopid = '" & replace(request.cookies("shopid"),"temp-","") & "'"
set rs = con.execute(stmt)
sn = rs("companyname")
set rs = nothing
%>
<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Untitled 1</title>
<style type="text/css">
.style1 {
	text-align: center;

}
.style2 {
	text-align: left;
	font-family: Arial, Helvetica, sans-serif;
}
.style3 {
	text-align: left;
	margin-left: 40px;
}
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

.style4 {
	text-align: center;
	color: #FFFFFF;
	background-image: url('newimages/pageheader.jpg');
	border-bottom: 2px black solid;
	font-family: Arial, Helvetica, sans-serif;
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
	font-family: Arial, Helvetica, sans-serif;
	color: #800000;
}

</style>
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

	if(document.bankform.paymethod.value == "mobank"){
		
		if(document.bankform.birthdate.value == ""){
			alert("Date of Birth is required for identity verification")
			return
		}
		if(document.bankform.last4.value == ""){
			alert("Last 4 of Social Security Number is required for identity verification")
			return
		}
		if(document.bankform.bankaccount.value == ""){
			alert("Bank Account Number is required")
			return
		}
		if(document.bankform.routing.value == ""){
			alert("Bank Routing Number is required")
			return
		}
		document.bankform.submit()
	}
	if(document.bankform.paymethod.value == "cc"){
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
		document.bankform.submit()
	}
}
</script>

</head>

<body>

<div style="text-align:center;height: 58px;background-image:url('newimages/rosubheader.jpg');font-family:Arial;">
<br/><span class="style6"><strong>To continue using Shop Boss Pro, please read 
	and accept our Subscription Agreement</strong></span><br/>

</div>

<p class="style1"><strong><span class="style5">Boss Software Inc.</span><br class="style5" />
<span class="style5">Subscription Agreement</span></strong></p>
<p class="style2">This agreement is made between Boss Software Inc. (&quot;Provider&quot;) 
and <%=sn%> (&quot;Subscriber&quot;) to provide access to the web-based application known as 
ShopBossPro.com (&quot;Service&quot;) in accordance with the Terms Of Service Agreement 
incorporated herein by reference.&nbsp; This Order Form (&quot;Order&quot;) shall define 
the terms under which this agreement is made.</p>
<p class="style2">1.&nbsp; Provider agrees to allow access to the Subscriber to 
its proprietary Service for the purposes of entering data, and creating Repair 
Orders.&nbsp; Repair Orders are defined as customer and vehicle information as 
well as customer repair requests entered into the Service.&nbsp; Once this data is 
entered, a Repair Order is generated indicating that a customer requested and/or 
authorized repair work to their vehicle.</p>
<p class="style2">2.&nbsp; Subscriber agrees to pay a month-to-month 
subscription fee to access the Service.&nbsp; This fee is based on Subscriber useage of the 
Service as described below</p>
<p class="style3"><span class="style5">
<strong>Silver Package </strong>- any calendar month the Subscriber creates 
1-30 Repair Orders in the Service, the monthly subscription fee will be $29.95;</span><br class="style5" />
<br class="style5" />
<span class="style5">
<strong>Gold Package</strong> - any calendar month the Subscrsiber creates 30-55 
Repair Orders in the Service, the monthly subscription fee will be $69.95;</span><br class="style5" />
<br class="style5" />
<span class="style5">
<strong>Platinum Package</strong> - any calendar month the Subscriber creates 55 
or more Repair Orders in the Service, the monthly subscription fee will be $99.95;</span></p>
		<p class="style3"><span class="style5"><strong>Estimating Guide</strong> - this is an add-on module providing third-party labor times 
and parts prices.&nbsp; This module is an additional $39.95 per month in 
addition to any subscribtion costs listed.</span><br class="style5" />
		</p>
<p class="style5">3.&nbsp; Cancellation:&nbsp; This agreement may be cancelled by either party 
for any reason with 30 days written notice.&nbsp; Upon cancellation, Subscriber 
will be offered the data entered into the Service by the Subscriber in 
comma-delimited format (CSV).&nbsp; Provider will maintain a backup copy for 12 
months after cancellation.</p>
<%
stmt = "select * from company where shopid = '" & replace(request.cookies("shopid"),"temp-","") & "'"
set rs = con.execute(stmt)

%>
<p>
<table cellpadding="6" cellspacing="0" style="width: 100%">
	<tr>
		<td class="style4" style="width: 50%; height: 53px;border-left:0px">
		<strong>Enter the information below and click on Accept Agreement</strong></td>
	</tr>
	<tr>
		<td style="width: 50%; ">
<form action="acceptagreement.asp" name="bankform" method="post">
<input name="dateofacceptance" value="<%=date%>" type="hidden" />
<input name="howpaying" value="mobank" type="hidden" />

	<table style="width: 100%" cellpadding="3" cellspacing="0">
		<tr>
			<td class="style5" style="width: 537px">I will be paying with
			<span class="style7"><strong>(Select Automatic Bank Debit or Credit 
			Card)</strong></span></td>
			<td>
			<select name="paymethod">
			<option <%=ccs%> value="cc">Automatic Credit Card Charge</option>
			</select></td>
		</tr>
		<tr style="<%=mhide%>" id="ba3">
			<td class="style5" style="width: 537px">
			My Date of Birth:</td>
			<td>
<input name="birthdate" value="<%=rs("birthdate")%>" type="text" class="style5" /></td>
		</tr>
		<tr style="<%=mhide%>" id="ba4">
			<td class="style5" style="width: 537px">Last 4 Digits of my Social Security Number:</td>
			<td>
<input name="last4" value="<%=rs("last4")%>" type="text" class="style5" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc0">
			<td class="style5" style="width: 537px"><b>Visa/Mastercard/Discover (NO AMEX)</b></td>
			<td class="style5"></td>
		</tr>
		<tr style="<%=cchide%>" id="cc1">
			<td class="style5" style="width: 537px">Credit Card Number (XXXX-XXXX-XXXX-XXXX)</td>
			<td class="style5">
<input name="cc" value='<%=rs("cc")%>'  type="text" class="style5" style="width: 175px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc2">
			<td class="style5" style="width: 537px">Expiration Date (MM/YY)</td>
			<td class="style5">
<input name="expdate" value='<%=rs("expdate")%>'  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc3">
			<td class="style5" style="width: 537px">Billing Address</td>
			<td class="style5">
<input name="billingaddress" value='<%=rs("billingaddress")%>'  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc4">
			<td class="style5" style="width: 537px">Billing City</td>
			<td class="style5">
<input name="billingcity" value='<%=rs("billingcity")%>'  type="text" class="style5" style="width: 264px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc5">
			<td class="style5" style="width: 537px">Billing State</td>
			<td class="style5">
<input name="billingstate" value='<%=rs("billingstate")%>'  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		<tr style="<%=cchide%>" id="cc6">
			<td class="style5" style="width: 537px">Billing Zip</td>
			<td class="style5">
<input name="billingzip" value='<%=rs("billingzip")%>'  type="text" class="style5" style="width: 104px" /></td>
		</tr>
		<tr>
			<td class="style5" style="width: 537px">&nbsp;</td>
			<td class="style5">
			&nbsp;</td>
		</tr>
		<tr>
			<td class="style5" style="width: 537px">Labor and Parts Guide</td>
			<td class="style5">
			<select name="lg">
			<option selected="selected" value="Yes">Include the Labor &amp; Parts Guide for additional $29.95 per month
			</option>
			<option value="No">Do not include the Labor &amp; Parts Guide
			</option>
			</select>
			</td>
		</tr>
	</table>
<p><span class="style5">Date of Acceptance by Subscriber: </span> <br />
<%
if len(rs("dateofacceptance")) > 0 then
	response.write rs("dateofacceptance")
else
	response.write date
end if

%>
<br />
<span class="style8"><strong>By clicking Accept Agreement below, I understand that I have read and understand the 
<a href="termsofuse.asp" target="_blank">Terms of Use</a> and I am authorizing Boss Software Inc. to automatically charge my credit card or bank account each month for my usage of Shop Boss Pro and the associated Labor Time and Parts Price Guide
</strong></span></p>
	<p>


<input name="Submit1" type="button" onclick="checkForm()" value="Accept Agreement" /></p>
		</form></td>
	</tr>
</table>

	</p>
	<p><br />
</p>

</body>

</html>
