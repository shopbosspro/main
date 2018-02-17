<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->

<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/main.css">
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
.style1 {
	font-size: 14px;
	color: white;
	background-image: url('newimages/pageheader.jpg');
	font-weight: bold;
	text-align: right;
}
.style12 {
	font-size: 14px;
	color: white;
	background-image: url('newimages/wipheader.jpg');
	font-weight: bold;
	text-align: right;
}

.tbl2header{
	font-size:14px;
	color:white;
	background-image:url('newimages/wipheader.jpg');
	font-weight:bold
}

</style>
</head>
<body >


 <table border="0" cellpadding="0" cellspacing="0" style="width: 80%" align="center">
  <tr>
   <td ondblclick="location.href='reports/statementpdf.asp?shopid=<%=request.cookies("shopid")%>'" class="sbheader">

   Account Tracking Module<br>
	
	<input onclick="location.href='company.asp'" class="obutton" type="button" value="Back" name="Abutton1">
	<input type="button" onclick="location.href='printaccounttracking.asp'" value="Print List" class="obutton" name="Abutton2">
	<!--<input type="button" onclick="location.href='pdfstatements/statementpdf.asp?shopid=<%=request.cookies("shopid")%>'" class="obutton" value="Print Stmts" name="Abutton2">
	-->
	<input type="button" onclick="location.href='pdfstatements/statementpdf.asp?shopid=<%=request.cookies("shopid")%>'" class="obutton" value="Print Stmts" name="Abutton2">

	</td>
  </tr>
  <tr>
   <td width="100%" valign="top">
   <p><strong>The following invoices are on account and have not been paid 
   <%	if date > "10/31/2012" then	%>
   (Red 
   header=RO's, Blue header=Parts Sales):</strong><br>
   <%	end if	%>
   <table style="width: 100%" cellspacing="0" cellpadding="3">
	<tr>
		<td class="tblheader" >RO #</td>
		<td class="tblheader" >Date Complete</td>
		<td class="tblheader">Days</td>
		<td class="tblheader">Customer Home / Work / Cell</td>
		<td style="text-align:right" class="tblheader">Total RO</td>
		<td style="text-align:right" class="tblheader">Balance</td>
		<td class="style1">Edit Account</td>
	</tr>
	<%
	function fphone(p)
		if len(p) > 0 then
			fphone = left(p,3) & "-" & mid(p,4,3) & "-" & right(p,4)
		else
			fphone = "NA"
		end if
	end function
	
	
	'********* get ro's with balances *************
	stmt = "select discountamt,totalfees,roid,finaldate,lastfirst,customerphone,cellphone,customerwork,totalro,balance from repairorders where shopid = '" _
	& request.cookies("shopid") & "' and round(balance,2) > 0 and (ucase(status) ='FINAL' or ucase(status) = 'CLOSED')" _
	& " order by lastfirst"
	set rs = con.execute(stmt)
	if not rs.eof then
		rs.movefirst
		do until rs.eof
		bal = bal + rs("balance")
		if len(rs("customerphone")) > 0 or len(rs("customerwork")) > 0 or len(rs("cellphone")) > 0 then
			lf = rs("lastfirst") & "<BR>" & fphone(rs("customerphone")) & " / " & fphone(rs("customerwork")) & " / " & fphone(rs("cellphone"))
		else
			lf = rs("lastfirst")
		end if
	%>
	<tr>
		<td style="width: 5%"><%=rs("roid")%>&nbsp;</td>
		<td style="width: 20%"><%=rs("finaldate")%>&nbsp;</td>
		<td style="width: 5%"><%=datediff("d",rs("finaldate"),date)%>&nbsp;</td>
		<td style="width: 30%"><%=lf%>&nbsp;</td>
		<td style="width: 15%;text-align:right"><%=formatcurrency(rs("totalro"))%>&nbsp;</td>
		<td style="width: 15%;text-align:right"><%=formatcurrency(rs("balance"))%>&nbsp;</td>
		<td style="width: 10%;text-align:right"><a target="_top" href='editaccount.asp?roid=<%=rs("roid")%>'>
		<img src="newimages/edit2.png" width="16" height="18"></a></td>
	</tr>
	<%
			rs.movenext
		loop
	else
		response.write "</table>No accounts to be collected"
	end if
	
	' now get outstanding parts sales
	stmt = "select * from ps where psdate > '2012-10-31' and round(balance,2) > 0 and shopid = '" & request.cookies("shopid") & "'"
	set rs = con.execute(stmt)
	if not rs.eof then
	%>
	<tr>
		<td class="tbl2header" >Invoice#</td>
		<td class="tbl2header" >Date Sold</td>
		<td class="tbl2header">Days</td>
		<td class="tbl2header">Customer</td>
		<td style="text-align:right" class="tbl2header">Total Invoice</td>
		<td style="text-align:right" class="tbl2header">Balance</td>
		<td class="style12">Edit Account</td>
	</tr>
	<%
		do until rs.eof
			stmt = "select * from customer where customerid = " & rs("cid") & " and shopid = '" & request.cookies("shopid") & "'"
			set trs = con.execute(stmt)
			if not trs.eof then
				lf = trs("lastname") & "," & trs("firstname")
			end if
			set trs = nothing
			bal = bal + rs("balance")
	%>
	<tr>
		<td style="width: 5%"><%=rs("psid")%>&nbsp;</td>
		<td style="width: 20%"><%=rs("psdate")%>&nbsp;</td>
		<td style="width: 5%"><%=datediff("d",rs("psdate"),date)%>&nbsp;</td>
		<td style="width: 30%"><%=lf%>&nbsp;</td>
		<td style="width: 15%;text-align:right"><%=formatcurrency(rs("total"))%>&nbsp;</td>
		<td style="width: 15%;text-align:right"><%=formatcurrency(rs("balance"))%>&nbsp;</td>
		<td style="width: 10%;text-align:right"><a target="_top" href='editpsaccount.asp?psid=<%=rs("psid")%>'>
		<img src="newimages/edit2.png" width="16" height="18"></a></td>
	</tr>
	<%
			rs.movenext
		loop
	end if
	%>

	</table>
	</td>
  </tr>
 </table>


<table style="width: 80%" cellspacing="0" cellpadding="3" align="center">
	<tr>
		<td  style="width: 5%">&nbsp;</td>
		<td  style="width: 20%">&nbsp;</td>
		<td  style="width: 5%">&nbsp;</td>
		<td  style="width: 25%">&nbsp;</td>
		<td style="text-align:right"><strong>Total Outstanding: &nbsp;&nbsp;&nbsp;&nbsp;<%=formatcurrency(bal)%></strong>&nbsp;</td>
		<td  style="width: 10%;text-align:right">&nbsp;</td>
	</tr>
	</table>
<script language="javascript">
	status = "Shop Boss Pro"

function closePrint(){

	document.getElementById("invoicediv").style.display = "none"
	document.getElementById("hider").style.display = "none"

}

//x = setTimeout("printRO()",200)

function printRO(){
	//alert("printing")
	document.getElementById("hider").style.display = "block"
	document.getElementById('timer').style.display = 'block'
	xmlHttp=GetXmlHttpObject()
	if (xmlHttp==null){
		alert ("Browser does not support HTTP Request")
		return
	} 
	var url="pdfstatements/statementpdf.asp?shopid=<%=request.cookies("shopid")%>"
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
		//document.getElementById("hider").style.display = "block"
		document.getElementById("invoicediv").style.display = "block"
		document.getElementById("invoice").src = "pdfstatements/temp/" + invname
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

</script>
<style type="text/css">
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

</style>
<div style="text-align:center;border:15px black solid;color:red;background-color:white;width:90%;height:90%;position:absolute;top:5%;left:5%;z-index:1050;display:none" id="invoicediv">
	<strong>Your Invoice is displayed below.&nbsp; Move your mouse 
	to the bottom to print it.&nbsp; Click here to</strong>
<input onclick="closePrint()" name="Button1" type="button" value="Close" >

<iframe frameborder="0" scrolling="no" style="width:95%;height:99%;" id="invoice" name="invoice"></iframe>
</div>
<div id="hider"></div>
<div id="timer">
Saving...<br><br>
<img src="newimages/ajax-loader.gif">
</div>

</html>
<%
'Copyright 2011 - Boss Software Inc.
%>