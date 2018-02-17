<!-- #include file=aspscripts/auth.asp --> <!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->



<%
stmt = "select usepartsmatrix from company where shopid = '" & request.cookies("shopid") & "'"
set rs = con.execute(stmt)
usematrix = lcase(rs("usepartsmatrix"))
set rs = nothing

	strsql = "Select * From partsregistry where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & request.querystring("PartNumber") & "'"
	'response.write strsql
	set rs = con.execute(strsql)
	if not rs.eof then 
		response.redirect "addpartfound.asp?partnum=" & request("partnum") & "&comid=" & request("comid") & "&roid=" & request("roid") & "&pid=" & rs("PartNumber")
	else
		'check for parts in inventory
		stmt = "select * from partsinventory where shopid = '" & request.cookies("shopid") & "' and partnumber = '" & request.querystring("partnumber") & "'"
		'response.write stmt
		set rs = con.execute(stmt)
		if not rs.eof then
			response.redirect "addpartfound.asp?comid=" & request("comid") & "&roid=" & request("roid") & "&pid=" & rs("PartNumber")
		end if
	
			
			
%>
                      <%
                      cstmt = "select * from category where shopid = '" & request.cookies("shopid") & "' order by Category, Start"
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	response.write "<script language='javascript'>"
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
<!-- #include file=javascripts/getfactor.js -->

<head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" src="javascripts/popwithdrop.js"></script>
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
.style1 {
	background-color: #FFFFFF;
}
.style3 {
	font-weight: bold;
}
.style4 {
	font-weight: bold;
		background-image: url('https://<%=request.servervariables("SERVER_NAME")%>/sbp/newimages/wipheader.jpg');
}
.style5 {
	color: #FFFFFF;
}

-->
</style>
<script language="javascript" src="javascripts/cookies.js"></script>
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
	}
}
</script>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<iframe id="newpo" style="width:60%;position:absolute;left:20%;top:150px;height:70%;background-color:white;display:none;z-index:1000"></iframe>
<div id="hider"></div>

<script type="text/javascript">
function addSupp(){

	document.getElementById("popup").style.display = "block"
	document.getElementById("popuphider").style.display = "block"
	document.getElementById("popupdropshadow").style.display = "block"
}

function closeSupplier(){

	document.getElementById("popup").style.display = "none"
	document.getElementById("popuphider").style.display = "none"
	document.getElementById("popupdropshadow").style.display = "none"
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
	if(document.theform.ReOrder.value > 0){
		document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;
	}
	if(document.theform.Price.value > 0){
		document.theform.TotalPrice.value = document.theform.Price.value*document.theform.Quantity.value;
	}
	if(document.theform.Discount.value >= 0){
		document.theform.Net.value = document.theform.Price.value-document.theform.Discount.value
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
	
	r = Math.random()
	document.getElementById("newpo").src = "addpartnewpo.asp?roid=<%=request("roid")%>&r=" + r
	document.getElementById("newpo").style.display = 'block'
	document.getElementById("popuphider").style.display = 'block'

}

</script>
<style type="text/css">
#popup{
	position:absolute;
	top:100px;
	left:100px;
	width:600px;
	height:500px;
	border:medium navy outset;
	text-align:center;
	color:black;
	display:none;
	z-index:999;
	background-color:white;
}
#popupdropshadow{
	position:absolute;
	top:115px;
	left:115px;
	width:600px;
	height:500px;
	text-align:center;
	color:black;
	display:none;
	z-index:998;
	background-color:black;
}
#popuphider{
	position:absolute;
	top:0px;
	left:0px;
	width:100%;
	height:300%;
	background-color:gray;
	-ms-filter:"progid:DXImageTransform.Microsoft.Alpha(Opacity=50)";
	filter: alpha(opacity=70); 
	-moz-opacity:.70; 
	opacity: .7;
	z-index:997;
	display:none;

}

.innerBox1 {
	font-weight: bold;
	background-image: url('newimages/wipheader.jpg');
	color: #FFFFFF;
	text-align:center;
}

.innerBox2 {
	background-color: #FFFFFF;
	text-align: center;
}
.style6 {
	color: #FF0000;
}
.style7 {
	color: #000000;
}
</style>
<div id="popuphider"></div>
<div style="display:none;" id="popup">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
        <tr>
          <td height="11" class="innerBox1"><strong>Enter your supplier data here</strong></td>
        </tr>
        </table>
   <form name=suppform action="addsuppliernow.asp" method=post>
   <input type=hidden name=comid value="<%=request("comid")%>">
   <input type="hidden" name="roid" value="<%=request("roid")%>">
   <input type="hidden" name="sp" value="addpart.asp">

   <table border="0" style="width: 100%" cellpadding="3" cellspacing="0">
    <tr>
     <td align="right" style="height: 26px; width: 42%;" class="">
	 <strong>Name</strong></td>
     <td width="77%" style="height: 26px" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierName" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Address</strong></td>
     <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierAddress" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>City</strong></td>
     <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierCity" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>State</strong></td>
     <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierState" size="7" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Zip</strong></td>
     <td width="77%" class="style2"><input type="text"  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="SupplierZip" size="11" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Phone</strong></td>
     <td width="77%" class=""><font size="3">(<input onkeyup="if(document.theform.sac.value.length == 3){document.theform.spre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">)
      <input type="text" onkeyup="if(document.theform.spre.value.length == 3){document.theform.sl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="spre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">-<input type="text" onkeyup="if(document.theform.sl4.value.length == 4){document.theform.sfac.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' name="sl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></font></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Fax</strong></td>
     <td width="77%" class=""><font size="3">(<input onkeyup="if(document.theform.sfac.value.length == 3){document.theform.sfpre.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfac" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">)
      <input onkeyup="if(document.theform.sfpre.value.length == 3){document.theform.sfl4.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfpre" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; ">-<input onkeyup="if(document.theform.sfl4.value.length == 4){document.theform.SupplierContact.focus()};return " onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="sfl4" size="5" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></font></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Contact</strong></td>
     <td width="77%" class="style2"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="SupplierContact" size="30" style="font-family: Arial; font-size: 10pt; font-variant: small-caps; "></td>
    </tr>
    <tr>
     <td align="right" class="" style="width: 42%"><strong>Active</strong></td>
     <td width="77%" class="style2"><select  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' size="1" name="Active">
       <option value="YES" selected>YES</option>
       <option value="NO">NO</option>
      </select></td>
    </tr>
    <tr>
     <td width="100%" colspan="2" class="innerBox2"><input onmouseover=this.style.cursor='hand' onclick="document.suppform.submit();" type="button" value="Add Supplier" name="B1">
	 <input name="Button2" onclick="closeSupplier()" type="button" value="Cancel"></td>
    </tr>
   </table>
   <input type="hidden" name="supid" value="<%=request("supid")%>">
   </form>

</div>
<div id="popupdropshadow"></div>


<div align="center">
 <center>
 <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
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
            <p align="center" style="height: 66px"><br><span class="style5">Add your parts info, then click on Add Part,</span><br class="style5">
			<span class="style5">or click on Save and Add Another to add multiple parts.</span></td>
          </tr>
         </table>
        </div>
        <div align="center">
         <center>
         <form method="post" name="theform" action="addpartnow.asp">
          <input type="hidden" name="rungetfactor" value="yes">
          <input type=hidden name=comid value="<%=request("comid")%>">
          <input type="hidden" name="roid" value="<%=request("roid")%>">
          <input type="hidden" name="s" value="addpart.asp">
          <table border="0" width="100%">
            <tr>
              <td width="100%"><table border="0" width="100%" cellspacing="0" cellpadding="2">
           <tr>
            <td width="30%" align="right" class="style3">Part Number:</td>
            <td width="35%"><input type="text" name="PartNumber" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("PartNumber")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Part Description:</td>
            <td width="35%"><input type="text" name="PartDesc" size="56" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("PartDesc")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Category:</td>
            <td width="35%"><select  size="1" name="Category">
                      <%
                      cstmt = "select distinct category from category where shopid = '" & request.cookies("shopid") & "'"
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	do while not crs.eof 
                      		response.write "<option value='" & crs("Category") & "'>" & crs("Category") & "</option>"
                      		crs.movenext
                      	loop
                      else
                      	response.write "<option value='General Part'>General Part</option>"
                      end if
                      %>
             </select></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Quantity:</td>
            <td width="35%"><input onblur="calcTotal()" type="text" name="Quantity" size="8" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("Quantity")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Cost:</td>
            <td width="35%"><input  onblur="calcTotal()" type="text" name="ReOrder" size="8" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("ReOrder")%>"></td>
            <td width="35%" class="style6"><strong>NOTE: To override matrix 
			pricing, change this to Yes BEFORE entering your prices for the part</strong></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3" style="height: 28px">Price:</td>
            <td width="35%" style="height: 28px"><input onblur="calcTotal()" type="text" name="Price" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("Price")%>"></td>
            <td width="35%" style="height: 28px">
            Override Matrix Price
			<select name="overridematrix">
			<%
			if usematrix = "no" then
				sselected = "selected='selected'"
				nselected = ""
			else
				sselected = ""
				nselected = "selected='selected'"
			end if
			%>
			<option <%=nselected%> value="No">No</option>
			<option <%=sselected%> value="Yes">Yes</option>
			</select>

            </td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Discount:</td>
            <td width="35%">
            <input onblur="if(this.value.length==0){this.value=0}" onkeyup="pd();" name="Discount" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="0"></td>
            <td width="35%"><div id="partd" style="visibility: hidden; width:124; height:16">
              <b><font color="#FF0000">Discount $0.00</font></b></div></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Net:</td>
            <td width="35%">
            <input onchange="javascript:document.theform.TotalPrice.value=document.theform.Net.value*document.theform.Quantity.value;document.theform.rungetfactor.value = 'no'" type="text" name="Net" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("net")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" style="height: 26px" class="style3">Code:</td>
            <td width="35%" style="height: 26px"><select  size="1" name="Codes">
                      <%
                      set crs = nothing
                      cstmt = "select * from codes where shopid = '" & request.cookies("shopid") & "' order by codes"
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	do while not crs.eof 
                      		response.write "<option value='" & crs("Codes") & "'>" & crs("Codes") & "</option>"
                      		crs.movenext
                      	loop
                      else
                      	response.write "<option value='No Codes'>No Codes</option>"
                      end if
                      %>
             </select></td>
            <td width="35%" style="height: 26px"></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Supplier:</td>
            <td width="35%" class="style1"><select  size="1" name="Supplier">
                      <%
                      set crs = nothing
                      cstmt = "select SupplierName from supplier where shopid = '" & request.cookies("shopid") & "' and length(SupplierName) > 1 and Active = 'YES' order by displayorder, SupplierName"
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	do while not crs.eof 
                      		response.write "<option value='" & crs("SupplierName") & "'>" & crs("SupplierName") & "</option>"
                      		crs.movenext
                      	loop
                      else
                      	response.write "<option value='None'>No Suppliers Entered</option>"
                      end if
                      %>
             </select>&nbsp; </td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6"><span class="style7">
			<strong>Purchase Order #:</strong></span> </td>
            <td width="35%" colspan="2"><span id="pospan">
            <%
            stmt = "select * from po where shopid = '" & request.cookies("shopid") & "' and saletype = 'RO' and salenumber = " & request("roid")
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
            <td width="18%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
           </tr>

           <tr>
            <td width="30%" align="right" class="style3">Bin: </td>
            <td width="35%" class="style1">
			<input type="text" name="bin" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value='<%=request("bin")%>'></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Invoice #:</td>
            <td width="35%" class="style1"><input type="text" name="InvoiceNo" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("InvoiceNo")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3">Extended Price:</td>
            <td width="35%" class="style1"><input type="text" name="TotalPrice" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("TotalPrice")%>"></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style3" style="height: 28px">Extended Cost:</td>
            <td width="35%" class="style1" style="height: 28px"><input name="MaxReOrder" size="11" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("MaxReOrder")%>"></td>
            <td width="35%" style="height: 28px"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Taxable:</strong> </td>
            <td width="35%" class="style1"><select name="tax">
            <%
            if len(request("tax")) > 0 then
            	response.write "<option value='" & request("tax") & "'>" & request("tax") & "</option>"
            end if
            %>
			<option value="Yes">Yes</option>
			<option value="No">No</option>
			</select></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Will This be an Inventoried 
			Part?</strong> </td>
            <td width="35%" class="style1"><select name="addtoinventory">
			<option value="No">No</option>
			<option value="Yes">Yes</option>
			</select></td>
            <td width="35%">&nbsp;</td>
           </tr>
           <tr>
            <td width="100%" colspan="3">
             <p align="center"><input type="button" value="Add Part" class="obutton" onclick="calcTotalSave();checkForm()" name="Abutton4">
				&nbsp;
				<input type="button" value="Save & Add Another" class="obutton" onclick="calcTotalSave();document.theform.addanother.value='Yes';checkForm()" name="Abutton3">
			 <input type="button" value="Cancel" class="obutton" onclick="window.event.returnValue = false;location.href='ro.asp?roid=<%=request("roid")%>'" name="Abutton5">
				</td>
           </tr>
          </table></td>
            </tr>
          </table><input type="hidden" name="sp" value="addpart.asp"><input type="hidden" name="addanother" value="No">
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

<p>
&nbsp;</p>


<script language="javascript">
document.theform.PartNumber.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
<%end if%>