<!-- #include file=aspscripts/auth.asp --> <!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=javascripts/getfactor.js -->


<%

if len(request("net")) > 0 then
	'connect to database and update information
	if len(request("discount")) > 0 then
		disamt = request("discount")
	else
		disamt = 0
	end if
	if len(request("net")) > 0 then
		netamt = request("net")
	else
		netamt = 0
	end if
	if len(request("comid")) > 0 then
		comid = request("comid")
	else
		comid = 0
	end if
	
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1
		
	stmt = "insert into parts (partid,shopid, discount, net, complaintid, PartNumber, PartDesc, PartPrice, Quantity, ROID, Supplier, bin, Cost,"
	stmt=stmt& " PartInvoiceNumber, PartCode, LineTTLPrice, LineTTLCost, `date`, PartCategory,tax, overridematrix,ponumber) values (" & newpartid & ",'" & request.cookies("shopid") & "', "
	stmt=stmt& disamt & ", "
	stmt=stmt& netamt & ", "
	stmt=stmt& comid & ", '"
	stmt=stmt& replace(request("PartNumber"),"'","''")
	stmt=stmt&"', '" & replace(request("PartDesc"),"'","''")
	if isnumeric(request("price")) then
		stmt = stmt & "', " & request("price")
	else
		stmt = stmt & "', 0"
	end if
	if isnumeric(request("quantity")) then
		stmt=stmt&", " & request("Quantity")
	else
		stmt=stmt&", 1"
	end if
	stmt=stmt&", " & request("roid")
	stmt=stmt&", '" & request("Supplier")
	stmt=stmt&"', '" & request("bin")
	stmt=stmt&"', " & request("ReOrder")
	stmt=stmt&", '" & replace(request("InvoiceNo"),"'","''")
	stmt=stmt&"', '" & request("Codes")
	stmt=stmt&"', " & request("TotalPrice")
	stmt=stmt&", " & request("MaxReOrder")
	stmt=stmt&", '" & dconv(Date())
	stmt=stmt&"', '" & request("Category")
	
	if len(request("ponumber")) > 0 then
		cstmt = "select * from po where id = " & request("ponumber") & " and shopid = '" & request.cookies("shopid") & "'"
		set rs = con.execute(cstmt)
		if not rs.eof then
			ponumber = cdbl(rs("ponumber"))
		else
			ponumber = 0
		end if
	else
		ponumber = 0
	end if
	stmt=stmt&"', '" & request("tax") & "', '" & request("overridematrix") & "', " & ponumber
	stmt=stmt&")"
	response.write stmt & "<BR>"
	stmt = ucase(stmt)
	con.execute stmt
	
	stmt = "Update partsregistry set PartDesc = '" & request("PartDesc")
	stmt = stmt & "', tax = '" & request("tax")
	stmt = stmt & "', PartPrice = " & request("Price")
	stmt = stmt & ", PartCode = '" & request("Codes")
	stmt = stmt & "', PartCost = " & request("ReOrder")
	stmt = stmt & ", PartSupplier = '" & ucase(replace(request("PartSupplier"),"'","''"))
	stmt = stmt & "', bin = '" & request("bin")
	stmt = stmt & "', PartCategory = '" & request("Category")
	stmt = stmt & "', overridematrix = '" & request("overridematrix")
	stmt = stmt & "' where PartNumber = '" & request("PartNumber") & "' and shopid = '" & request.cookies("shopid") & "'"
	'response.write stmt & "<BR>"
	stmt = ucase(stmt)
	con.execute(stmt)
	
	' now set inventory levels
	stmt = "update partsinventory set onhand = " & request("adjustedonhand") & ", netonhand = " & request("adjustedonhand") & " where shopid = '" & request.cookies("shopid") & "' and partnumber = '" & request("PartNumber") & "'"
	con.execute(stmt)
	
	if request("addanother") = "Yes" then
		response.redirect "addpart1.asp?comid=" & request("comid") & "&roid=" & request("roid")
	else
		response.redirect "ro.asp?ap=yes&roid=" & request("roid")
	end if


end if


stmt = "select usepartsmatrix from company where shopid = '" & request.cookies("shopid") & "'"
set rs = con.execute(stmt)
usepartsmatrix = rs("usepartsmatrix")
set rs = nothing

strsql = "Select * From partsregistry where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & request("pid") & "'"
set rs = con.execute(strsql)
if rs.eof then 
	response.redirect "addpart.asp?roid?=" & request("roid")
else
	pnum = rs("PartNumber")
	
	' get qtys from inventory
	stmt = "select onhand,allocatted,netonhand from partsinventory where shopid = '" & request.cookies("shopid") & "' and partnumber = '" & pnum & "'"
	set trs = con.execute(stmt)
	if not trs.eof then
		onhand = trs("onhand")
		allocated = trs("allocatted")
		netonhand = trs("netonhand")
	else
		onhand = rs("onhand")
		allocated = rs("allocatted")
		netonhand = rs("netonhand")
	end if
	rungetfactor = "no"
	cstmt = "select shopid,category,factor,start,end from category where shopid = '" & request.cookies("shopid") & "' order by Category, Start"
	'response.write cstmt
	set crs = con.execute(cstmt)
	if not crs.eof then
	rungetfactor = "yes"
	crs.movefirst
	response.write "<script language='javascript'>" & chr(10)
	cntr = 0
	response.write "var clist = new Array()" & chr(10)
	do while not crs.eof 
		response.write "clist["&cntr&"] = {Category:'" & lcase(crs("Category")) & "', Start:" & crs("Start") & ", End:" & crs("End") & ", Factor:" & crs("Factor") & "}" & chr(10)
		crs.movenext
		cntr = cntr + 1
	loop
end if
response.write "</script>"
set pcrs = nothing
set crs = nothing
%>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script language="Javscript" type="text/javascript" src="../javascripts/killenter.js"></script>
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--
.menubutton{
	width:120px;
	height:66px;
	font-family:Arial, Helvetica, sans-serif;
	font-size:small;
	font-weight:bold;
	color:maroon;
	cursor:hand;
}
.menubutton:hover{
	color:blue;
}
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
.style3 {
	border: 1px solid #000000;
}
.style4 {
	font-weight: bold;
		background-image: url('https://<%=request.servervariables("SERVER_NAME")%>/sbp/newimages/wipheader.jpg');
}
.style5 {
	color: #FFFFFF;
}
.style6 {
	font-weight: bold;
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

.auto-style1 {
	text-align: center;
	border: 4px solid #FF0000;
}
.auto-style2 {
	text-align: center;
	color: #FFFFFF;
	background-image:url("newimages/pageheader.jpg")
}
.auto-style3 {
	text-align: center;
}

-->
</style>
<script language="javascript">
function pd(){
	if (document.theform.Discount.value >= 0){
		mypd = (document.theform.Price.value * document.theform.Discount.value) / 100
		mypd = Math.round(mypd*100)/100
		mynetpd = Math.round((document.theform.Price.value - mypd)*100)/100
		document.theform.Net.value = mynetpd
		partd.innerHTML = "<b><font color='#FF0000'>Discount: $" + mypd + "</font></b>"
		partd.style.visibility = "visible"
		document.theform.TotalPrice.value = document.theform.Net.value*document.theform.Quantity.value
	}else{
		if(document.theform.Discount.value.length == 0){
			document.theform.Discount.value = 0
		}
	}
}
function isNumber(n) { 
  return !isNaN(parseFloat(n)) && isFinite(n); 
} 

function checkForm(){
	d = document.theform
	if(d.PartNumber.value.length == 0){
		alert("Part Number is required")
		return false;
	}
	if(d.PartDesc.value.length == 0){
		alert("Part Description is required")
		return false;
	}
	if(!isNumber(d.ReOrder.value)){
		alert("Cost is required and must be a number without a dollar signs")
		return false;
	}
	if(!isNumber(d.Quantity.value)){
		alert("Quantity is required and must be a number")
		return false;
	}
	if(!isNumber(d.Discount.value)){
		alert("Discount must be 0 (zero) or a number reflecting a desired discount")
		return false;
	}
	if(!isNumber(d.Price.value)){
		alert("Price is required and must be a number without a dollar sign")
		return false;
	}
	if(!isNumber(d.MaxReOrder.value)){
		alert("Total Cost is required and must be a number without a dollar sign")
		return false;
	}
	if(!isNumber(d.Net.value)){
		alert("Net Price is required and must be a number without a dollar sign")
		return false;
	}
	if(!isNumber(d.TotalPrice.value)){
		alert("Total Price is required and must be a number without a dollar sign")
		return false;
	}


	document.theform.submit()
}

function calcTotal(){
	if(document.theform.Discount.value >= 0){
		document.theform.Net.value = document.theform.Price.value*(1-(document.theform.Discount.value/100))
	}
	if(document.theform.ReOrder.value > 0){
		document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;
	}
	if(document.theform.Price.value > 0){
		document.theform.TotalPrice.value = document.theform.Net.value*document.theform.Quantity.value;
	}
	if(!isNumber(document.theform.currentonhand.value)){
		document.theform.currentonhand.value = 0
	}
	if(!isNumber(document.theform.adjustedonhand.value)){
		document.theform.adjustedonhand.value = 0
	}

	//alert(document.theform.overridematrix.value)
	if (document.theform.overridematrix.value != 'Yes'){
		getFactor()
	}
}

function calcTotalSave(){
	
	calcTotal()
	document.theform.TotalPrice.value=document.theform.Net.value*document.theform.Quantity.value;
	document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;

}

function createNewPO(){

	document.getElementById("newpo").src = "addpartfoundnewpo.asp?roid=<%=request("roid")%>"
	document.getElementById("newpo").style.display = 'block'
	document.getElementById("hider").style.display = 'block'

}


function calcInv(qty){
	
	qonhand = document.theform.currentonhand.value
	
	if(!isNumber(qonhand)){qonhand = 0}
	
	if(qonhand > 0){
		document.theform.adjustedonhand.value = qonhand - qty;
	}

}
</script>
</head>
<%
if rs("overridematrix") = "No" then
	onload = "onload='pd();calcInv(1)'"
else
	onload = "onload='calcInv(1)'"
end if
%>
<body  <%=onload%> link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<iframe id="newpo" style="width:60%;position:absolute;left:20%;top:150px;height:70%;background-color:white;display:none;z-index:1000"></iframe>
<div id="hider"></div>

<div align="center">
 <center>
 <table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
   <td width="100%" valign="top">
    <div align="left">
     <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
      <tr>
       <td valign="top" width="100%">
        <div align="left">
         <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
           <td height="11" class="style4">
            <p align="center" style="height: 52px">
			<br><span class="style5">Complete any missing or updated information,
            then click on Add Part</span></td>
          </tr>
         </table>
        </div>
        <div align="center">
         <center>
         <form method="post" action="addpartfound-new.asp" name="theform">
          <table width="80%" cellspacing="0" cellpadding="3" class="style3">
           <tr>
            <td width="30%" align="right" class="style6">Part Number:</td>
            <td width="35%" colspan="2"><input onkeyup="" type="text" name="PartNumber" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartNumber")%>"></td>
            <td width="35%" colspan="2">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Part Description:</td>
            <td width="35%" colspan="2"><input onkeyup="" type="text" name="PartDesc" size="56" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=replace(rs("PartDesc"),"""","&quot;")%>"></td>
            <td width="35%" colspan="2">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6" style="height: 28px">Category:</td>
            <td width="35%" style="height: 28px" colspan="2"><select onkeyup="" size="1" name="Category">
                      <option selected value="<%=rs("PartCategory")%>"><%=rs("PartCategory")%></option>
                      
                      &nbsp;<%
                      cstmt = "select distinct category from category where shopid = '" & request.cookies("shopid") & "' order by displayorder"
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	do while not crs.eof 
                      		if lcase(request.cookies("matrix")) = "no" and crs("category") = "NON MATRIX" then
                      			s = " selected='selected' "
                      		else
                      			s = ""
                      		end if
                      		response.write "<option " & s & " value='" & crs("Category") & "'>" & crs("Category") & "</option>"
                      		crs.movenext
                      	loop
                      else
                      	response.write "<option value='General Part'>General Part</option>"
                      end if
                      %>
                                                                  &nbsp;&nbsp;&nbsp;&nbsp; </select></td>
            <td width="35%" colspan="2" style="height: 28px"></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Quantity:</td>
            <td width="35%" colspan="2"><input onblur="calcTotal();calcInv(this.value)" onkeyup="" type="text" name="Quantity" size="8" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="1">
            <%
            'get inv qty
            mstmt = "select NetOnHand from partsinventory where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & pnum & "'"
            set mrs = con.execute(mstmt)
            if not mrs.eof then
            	mrs.movefirst
            	mqty = mrs("NetOnHand")
            else
            	mqty = 0
            end if
            set mrs = nothing
            response.write "(Qty in stock - " & mqty & ")"
            'get parts sales history
            pstmt = "select * from parts where deleted = 'no' and shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & pnum & "'"
            set prs = con.execute(pstmt)
            if not prs.eof then
            	prs.movefirst
            	while not prs.eof
            	pcost = prs("cost")
            	pprice = prs("partprice")
            	proid = prs("roid")
            	prs.movenext
            	wend
            end if
            if len(proid) > 0 then
            pstmt = "Select DateIn from repairorders where shopid = '" & request.cookies("shopid") & "' and roid = " & proid 
            set prs = con.execute(pstmt)
            if not prs.eof then
            	prs.movefirst
            	dsold = prs("datein")
            end if
            set prs = nothing
            end if
            'pstmt = "select roid from parts where shopid = '" & request.cookies("shopid") & "' and partnumber = '" pnum & '"
            'set prs = con.execute
            'if not prs.eof then
            '	prs.movefirst
            	
            	
            %>
            </td>
            <td width="35%" colspan="2">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Cost:</td>
            <td width="35%" style="width: 0%">
			<input id="cost" onkeyup="" onblur="calcTotal()" name="ReOrder" size="11" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartCost")%>">
			</td>
            <td width="35%" style="width: 17%"> &nbsp;</td>
            <td width="35%" colspan="2">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Price:</td>
            <td width="35%" colspan="2">
            <input onblur="calcTotal()" onkeyup="" type="text" name="Price" size="13" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartPrice")%>"></td>
            <td width="35%" colspan="2">
            Override Matrix Price
			<select name="overridematrix">
			<%
			if lcase(rs("overridematrix")) = "yes" then
				yselect = "selected='selected'"
			else
				yselect = ""
			end if
			if lcase(rs("overridematrix")) = "no" then
				nselect = "selected='selected'"
			else
				nselect = ""
			end if
			if lcase(usepartsmatrix) = "no" then
				yselect = "selected='selected'"
				nselect = ""
			end if
			%>
			<option <%=nselect%> value="No">No</option>
			<option <%=yselect%> value="Yes">Yes</option>
			</select>
			</td>
           </tr>
           <tr>
            <td width="30%" align="right" valign="top" class="style6">Discount %:</td>
            <td width="35%" valign="top" colspan="2">
            <input onblur="if(this.value.length==0){this.value=0}" onkeyup="pd();" type="text" name="Discount" size="13" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%if rs("discount") > 0 then response.write rs("discount") else response.write "0" end if%>"></td>
            <td width="35%" colspan="2" valign="top"><div id="partd" style="visibility: hidden; width:124; height:16">
              <b><font color="#FF0000">Discount $0.00</font></b></div></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6" style="height: 26px">Net Price:</td>
            <td width="35%" style="height: 26px" colspan="2">
            <%
            if rs("net") > 0 then
            	thenet = rs("net")
            else
            	thenet = rs("partprice")
            end if
            %>
            <input onkeyup="" type="text" name="Net" size="13" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=thenet%>"></td>
            <td width="35%" colspan="2" style="height: 26px"></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Code:</td>
            <td width="35%" colspan="2"><select onkeyup="" size="1" name="Codes">
              <option selected value="<%=rs("PartCode")%>"><%=rs("PartCode")%></option>
             </select></td>
            <td width="18%">Last Sold on </td>
            <td width="17%"><%=dsold%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Supplier:</td>
            <td width="35%" colspan="2"><select size="1" onkeyup="" name="Supplier">
              <option selected value="<%=rs("PartSupplier")%>"><%=rs("PartSupplier")%></option>
              <%
              stmt = "select * from supplier where shopid = '" & request.cookies("shopid") & "' order by displayorder,suppliername"
              set suprs = con.execute(stmt)
              if not suprs.eof then
              	do until suprs.eof
              		response.write "<option value='"
              		if rs("partsupplier") = suprs("suppliername") then response.write "selected='selected'"
              		response.write suprs("suppliername") & "'>" & suprs("suppliername") & "</option>"
              		suprs.movenext
              	loop
              end if
              %>
&nbsp; 
             </select>&nbsp;
			</td>
            <td width="18%">Last Cost</td>
            <td width="17%"><%=pcost%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Purchase Order #: </td>
            <td width="35%" colspan="2"><span id="pospan">
            <%
            stmt = "select * from po where shopid = '" & request.cookies("shopid") & "' and saletype = 'RO' and salenumber = " & request("roid")
            'response.write stmt
            set ars = con.execute(stmt)
            if not ars.eof then
            	
            	response.write "<select name='ponumber' id='ponumber'>"
            	do until ars.eof
            		if cint(ars("ponumber")) = cint(request("newpo")) then
            			s = "selected='selected'"
            		else
            			s = ""
            		end if
            		response.write "<option " & s & " value='" & ars("id") & "'>" & ars("ponumber") & " - " & ars("issuedto") & "</option>"
            		ars.movenext
            	loop
            end if
            %></span>
            <input onclick="createNewPO()" style="font-size:x-small;margin-left:5px; width: 71px;" name="Button1" type="button" value="New PO" >
            </td>
            <td width="18%">Price Last Sold</td>
            <td width="17%"><%=formatcurrency(pprice)%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Bin: </td>
            <td width="35%" colspan="2">
			<input onkeyup="" type="text" name="bin" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value='<%=rs("bin")%>'></td>
            <td width="18%">Sold Prev. 6 Mos.</td>
            <td width="17%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Invoice #:</td>
            <td width="35%" colspan="2"><input onkeyup="" type="text" name="InvoiceNo" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value=""></td>
            <td width="18%" class="auto-style1" colspan="2" rowspan="4" style="width: 35%" valign="top">
			<strong>PLEASE VERIFY INVENTORY ADJUSTMENTS HERE.&nbsp; PLEASE NOTE, 
			NO INVENTORY ADJUSTMENTS WILL BE DONE WHEN CLOSING THE REPAIR ORDER.&nbsp; 
			CHANGES MUST BE MADE HERE.
			<br></strong>
			<table cellpadding="3" cellspacing="0" style="width: 100%">
				<tr>
					<td class="auto-style2" style="width: 50%"><strong>Current 
					On Hand</strong></td>
					<td class="auto-style2" style="width: 50%"><strong>On Hand 
					after adding this part to Repair Order</strong></td>
				</tr>
				<tr>
					<td class="auto-style3" style="width: 50%">
					<input tabindex="-1" name="currentonhand" style="width: 73px" type="text" value="<%=onhand%>"></td>
					<td class="auto-style3" style="width: 50%">
					<input tabindex="-1" name="adjustedonhand" style="width: 73px" type="text" value="<%=onhand%>"></td>
				</tr>
			</table>
			   </td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Total Price:</td>
            <td width="35%" colspan="2"><input onkeyup="" onfocus="javascript:this.value=document.theform.Net.value*document.theform.Quantity.value;" type="text" name="TotalPrice" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value=""></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Total Cost:</td>
            <td width="35%" style="width: 0%">
			<input id="tcost" name="MaxReOrder" size="11" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="">
			<input id="cost1" onkeyup="" onblur="document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;getFactor()" name="cost0" size="11" style="display:none; font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartCost")%>"> </td>
            <td width="35%" style="width: 17%">
			&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Taxable:</strong> </td>
            <td width="35%" colspan="2"><select name="tax">
            <%
            	response.write "<option value='" & rs("tax") & "'>" & rs("tax") & "</option>"
            %>
			<option value="Yes">Yes</option>
			<option value="No">No</option>
			</select></td>
           </tr>
           <tr>
            <td width="100%" colspan="5">
             <p align="center"><input value="Add Part" onclick="calcTotalSave();checkForm()" class="obutton" type="button" name="Abutton5">
				
				<input type="button" value="Save & Add Another" class="obutton" onclick="calcTotalSave();document.theform.addanother.value='Yes';checkForm()" name="Abutton4">
				
				<input type="button" value="Search Again" class="obutton" onclick='location.href="addpart1.asp?comid=<%=request("comid")%>&roid=<%=request("roid")%>"' name="Abutton1">
				
            </td>
           </tr>
          </table>
          <input type="hidden" name="pid" value="<%=request("pid")%>"><input type="hidden" name="roid" value="<%=request("roid")%>"><input id="rungetfactor" type="hidden" name="rungetfactor" value="no"><input type="hidden" name="sp" value="addpartfound.asp"><input type="hidden" name="addanother" value="No">
          <input type="hidden" name="comid" value="<%=request("comid")%>">
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
<p>&nbsp;</p>

<script language="javascript">
document.theform.PartNumber.focus();
document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;
<%
if rungetfactor = "yes" then response.write "document.getElementById('rungetfactor').value = 'yes'" & chr(10)
%>
if (document.theform.overridematrix.value != 'Yes'){
	getFactor()
}
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
<%end if%>