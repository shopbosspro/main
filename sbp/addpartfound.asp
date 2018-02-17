<!-- #include file=aspscripts/auth.asp --> <!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=javascripts/getfactor.js -->


<%
shopid = request.cookies("shopid")
stmt = "SELECT table_name FROM information_schema.tables WHERE table_schema = 'shopboss' AND table_name = 'partsregistry-" & shopid & "' LIMIT 1"
'response.write stmt & "<BR>"
set rs = con.execute(stmt)
'response.write rs.eof & "<BR>"
if not rs.eof then
	preg = "`partsregistry-" & shopid & "`"
else
	preg = "partsregistry"
end if

stmt = "select usepartsmatrix from company where shopid = '" & request("shopid") & "'"
set rs = con.execute(stmt)
usepartsmatrix = rs("usepartsmatrix")
set rs = nothing

strsql = "Select * From " & preg & " where shopid = '" & request("shopid") & "' and PartNumber = '" & request("pid") & "'"
set rs = con.execute(strsql)
if rs.eof then 
	response.redirect "addpart.asp?roid?=" & request("roid")
else
	pnum = rs("PartNumber")
	rungetfactor = "no"
	cstmt = "select shopid,category,factor,start,end from category where shopid = '" & request("shopid") & "' order by Category, Start"
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
}
.auto-style2 {
	color: #FF0000;
}

-->
</style>
<script type="text/javascript" src="javascripts/jquery-1.8.0.min.js"></script>
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

function checkBox(id){

	if(id == "useexisting"){
		document.getElementById("usenew").checked = false;
		document.getElementById('quantityonseparateorder').value = ""
	}else{
		document.getElementById("useexisting").checked = false;
		document.getElementById('quantityonseparateorder').value=document.getElementById('Quantity').value
	}

}
</script>
</head>
<%
if rs("overridematrix") = "No" then
	onload = "onload='pd()'"
else
	onload = ""
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
         <form method="post" action="addpartnow.asp" name="theform" id="theform">
          <table width="80%" cellspacing="0" cellpadding="3" class="style3">
           <tr>
            <td width="30%" align="right" class="style6">Part Number:</td>
            <td width="35%" colspan="2"><input onkeyup="" type="text" name="PartNumber" size="22" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartNumber")%>"></td>
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
                      cstmt = "select distinct category from category where shopid = '" & request("shopid") & "' order by displayorder"
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
            <td width="35%" colspan="2"><input onblur="calcTotal();" onkeyup="" type="text" id="Quantity" name="Quantity" size="8" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="1">
            <%
            'get inv qty
            mstmt = "select NetOnHand,partcost,partprice from partsinventory where shopid = '" & request("shopid") & "' and PartNumber = '" & pnum & "'"
            set mrs = con.execute(mstmt)
            if not mrs.eof then
            	mrs.movefirst
            	mqty = mrs("NetOnHand")
            	iprice = mrs("partprice")
            	icost = mrs("partcost")
            else
            	iprice = rs("partprice")
            	icost = rs("partcost")
            	mqty = 0
            end if
            set mrs = nothing
            response.write "(Qty in stock - " & mqty & ")"
            'get parts sales history
            pstmt = "select * from parts where deleted = 'no' and shopid = '" & request("shopid") & "' and PartNumber = '" & pnum & "' order by ts desc"
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
            pstmt = "Select DateIn from repairorders where shopid = '" & request("shopid") & "' and roid = " & proid 
            set prs = con.execute(pstmt)
            if not prs.eof then
            	prs.movefirst
            	dsold = prs("datein")
            end if
            set prs = nothing
            end if
            'pstmt = "select roid from parts where shopid = '" & request("shopid") & "' and partnumber = '" pnum & '"
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
			<input id="cost" onkeyup="" onblur="calcTotal()" name="ReOrder" size="11" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=icost%>">
			</td>
            <td width="35%" style="width: 17%"> &nbsp;</td>
            <td width="35%" colspan="2">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Price:</td>
            <td width="35%" colspan="2">
            <input onblur="calcTotal()" onkeyup="" type="text" name="Price" size="13" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=iprice%>"></td>
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
              <%
              stmt = "select * from codes where shopid = '" & request("shopid") & "'"
              set pcrs = con.execute(stmt)
              if not pcrs.eof then
              	do until pcrs.eof
              %>
              <option value="<%=pcrs("codes")%>"><%=pcrs("codes")%></option>
              <%
              	pcrs.movenext
              	loop
              end if
              %>
             </select></td>
            <td width="18%">Last Sold on </td>
            <td width="17%"><%=dsold%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Supplier:</td>
            <td width="35%" colspan="2"><select size="1" onkeyup="" name="Supplier">
              <option selected value="<%=rs("PartSupplier")%>"><%=rs("PartSupplier")%></option>
              <%
              stmt = "select * from supplier where shopid = '" & request("shopid") & "' order by displayorder,suppliername"
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
            <td width="18%">Last Reorder</td>
            <td width="17%"><%=pcost%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Purchase Order #: </td>
            <td width="35%" colspan="2"><span id="pospan">
            <%
            stmt = "select * from po where shopid = '" & request("shopid") & "' and saletype = 'RO' and salenumber = " & request("roid")
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
            <td width="18%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Bin: </td>
            <td width="35%" colspan="2">
			<input onkeyup="" type="text" name="bin" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value='<%=rs("bin")%>'></td>
            <td width="18%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Invoice #:</td>
            <td width="35%" colspan="2"><input onkeyup="" type="text" name="InvoiceNo" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value=""></td>
            <td width="18%">Price Last Sold</td>
            <td width="17%"><%=formatcurrency(pprice)%></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Total Price:</td>
            <td width="35%" colspan="2"><input onkeyup="" onfocus="javascript:this.value=document.theform.Net.value*document.theform.Quantity.value;" type="text" name="TotalPrice" size="22" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value=""></td>
            <td width="18%">Sold Prev. 6 Mos.</td>
            <td width="17%">&nbsp;</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style6">Total Cost:</td>
            <td width="35%" style="width: 0%">
			<input id="tcost" name="MaxReOrder" size="11" style="font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="">
			<input id="cost1" onkeyup="" onblur="document.theform.MaxReOrder.value=document.theform.ReOrder.value*document.theform.Quantity.value;getFactor()" name="cost0" size="11" style="display:none; font-family: Arial; font-size: 8pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=rs("PartCost")%>"> </td>
            <td width="35%" style="width: 17%">
			&nbsp;</td>
            <td width="35%" colspan="2" rowspan="3">&nbsp;</td>
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
           <%
           'if request.cookies("shopid") = "1734" then
           'check for part in stock
           if mqty > 0 then
           %>
            <td width="30%" class="auto-style1" colspan="3" style="width: 65%;border:3px red outset">
			<span class="auto-style2"><strong>This Part has a quantity of&nbsp;<%=mqty%>
			in stock.<br>Should Shop Boss pull this part from inventory<br>or 
			are you adding it from a separate parts order?<br><br></strong></span><table style="width: 100%">
				<tr>
					<td valign="top"><input checked="checked" onclick="checkBox(this.id)" id="useexisting" value="yes" name="useexisting" type="checkbox"><span class="auto-style2"><strong>Add this part 
					from inventory</strong></span></td>
					<td valign="top"><input onclick="checkBox(this.id)" id="usenew" name="usenew" type="checkbox"><span class="auto-style2"><strong>Add this part from 
					separate order<br>
					<input name="quantityonseparateorder" id="quantityonseparateorder" style="width: 24px" type="text">Quantity 
					on separate order</strong></span></td>
				</tr>
			</table>
			   </td>
			<%
			end if
			'end if
			%>
           </tr>
           <tr>
            <td width="100%" colspan="5">
             <p align="center"><input id="addpartbutton" value="Add Part" onclick="calcTotalSave();checkForm()" class="obutton" type="submit" name="Abutton5">
				
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

$('theform').keypress( function( e ) {
  var code = e.keyCode || e.which;

  if( code === 13 ) {
    e.preventDefault();
    return false; 
  }
})

</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
<%end if%>