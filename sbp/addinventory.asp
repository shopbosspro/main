<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->

<!-- #include file=aspscripts/adovbs.asp -->
<%
cstmt = "select * from category where shopid = '" & request.cookies("shopid") & "' order by Category, Start"
set crs = con.execute(cstmt)
if not crs.eof then
	crs.movefirst
   	response.write chr(10) & "<script language='javascript'>" & chr(10)
   	cntr = 0
   	response.write "var clist = new Array()" & chr(10)
   	do while not crs.eof 
   		response.write "clist["&cntr&"] = {Category:'" & lcase(crs("Category")) & "', Start:" & crs("Start") & ", End:" & crs("End") & ", Factor:" & crs("Factor") & "}" & chr(10)
       crs.movenext
       cntr = cntr + 1
    loop
end if
response.write "</script>"
set crs = nothing
%>

<html>
<head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" src="javascripts/killenter.js"></script>
<script type="text/javascript" src="jquery/jquery-1.10.2.min.js"></script>
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--
.obutton{
	height:40px;
	width:80px;
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
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<script language="Javascript">
function getFactor(){
	//alert(document.theform.PartCost.value)
	srch = document.theform.PartCategory.value.toLowerCase()
	amt = document.theform.PartCost.value
	var i
	for (i=0; i<clist.length; i++){
		if (clist[i].Category.indexOf(srch) != -1 && clist[i].Start <= amt && clist[i].End >= amt){
		 	document.theform.PartPrice.value = Math.round((amt * clist[i].Factor)*100)/100
		}
	}
}

function checkDigits(v){

	for(j=0;j<v.length;j++){
		curchar = v.charCodeAt(j)
		curlett = v.charAt(j)
		if(curchar >= 45 && curchar <= 57 || curchar >= 65 && curchar <= 90 || curchar >= 97 && curchar <= 122 || curchar == 95){
			//console.log("ok:" + curchar)
		}else{
			alert("Only letters,numbers and - can be used in the Part Number field")
			newv = v.replace(curlett,"")
			document.theform.PartNumber.value = newv
		}
	}

}
	
function checkForm(){
	d = document.theform
	if(d.PartCategory.value.length == 0){
		alert("Category is required")
		return false;
	}
	if(d.PartNumber.value.length == 0){
		alert("Part Number is required")
		return false;
	}
	if(d.PartDesc.value.length == 0){
		alert("Description is required")
		return false;
	}
	if(d.PartCode.value.length == 0){
		alert("Code is required")
		return false;
	}
	if(d.PartPrice.value.length == 0){
		alert("Price is required")
		return false;
	}
	if(d.PartCost.value.length == 0){
		alert("cost is required")
		return false;
	}
	if(d.OnHand.value.length == 0){
		alert("On Hand is required")
		return false;
	}
	if(d.ReOrderLevel.value.length == 0){
		alert("Re-Order Level is required")
		return false;
	}
	if(d.MaxOnHand.value.length == 0){
		alert("Max On Hand is required")
		return false;
	}
	if(d.MaintainStock.value.length == 0){
		alert("Maintain Stock is required")
		return false;
	}
	if(d.tax.value.length == 0){
		alert("Taxable is required")
		return false;
	}


	document.theform.submit()
}

function checkPartNumber(pn){

	$.ajax({
		data: "shopid=<%=request.cookies("shopid")%>&pn=" + encodeURIComponent(pn),
		url: "aspscripts/checkpartnumber.asp",
		success: function(r){
			if (r == "found"){
				alert("This part number is already on file.  Please enter a new part number")
				document.theform.PartNumber.focus()
			}
		}
	
	});
	

}
</script>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

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
           <td style="height: 51px;background-image:url('newimages/wipheader.jpg')">
            <p align="center" class="style1">Add your parts info, then click on Add to
            Inventory.</td>
          </tr>
         </table>
        </div>
        <div align="center">
         <center>
         <form method="post" name="theform" action="addinventorynow.asp">
         <input type=hidden name=rungetfactor value="no">
         <table border="0" width="100%">
            <tr>
              <td width="65%">
			  <table border="0" width="100%" cellspacing="0" cellpadding="4">
           <tr>
            <td width="30%" align="right"><b>Part Number:</b></td>
            <td width="70%" colspan="2">
			<input onkeyup="checkDigits(this.value)" type="text" onblur="checkPartNumber(this.value)" name="PartNumber" size="22" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; tabindex="1" value='<%=request("partnumber")%>' tabindex="1"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Part Description:</b></td>
            <td width="70%" colspan="2"><input  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC'" type="text" name="PartDesc" size="56" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; value="<%=request("PartDesc")%>" tabindex="2"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Category:</b></td>
            <td width="70%" colspan="2"><select  size="1" name="PartCategory" tabindex="3">
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
           </tr>
           <tr>
            <td width="30%" align="right"><b>Cost:</b></td>
            <td width="70%" colspan="2">
			<input  onblur="getFactor();this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" name="PartCost" size="11" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; height: 20px;" tabindex="4" value="0"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Price:</b></td>
            <td width="70%" colspan="2">
			<input  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" type="text" name="PartPrice" size="11" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; value="<%=request("Price")%>" tabindex="5" value="0"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Code: </b></td>
            <td style="width: 79px"><select  size="1" name="PartCode" tabindex="6">
                      <%
                      set crs = nothing
                      cstmt = "select * from codes where shopid = '" & request.cookies("shopid") & "' order by Codes"
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
            <td width="70%" style="width: 35%">
            Override Matrix Price
			<select name="overridematrix" tabindex="7">
			<option value="No">No</option>
			<option value="Yes">Yes</option>
			</select>

            </td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Supplier: </b></td>
            <td width="70%" colspan="2">
			<select  size="1" name="PartSupplier" tabindex="8">
                      <%
                      set crs = nothing
                      cstmt = "select SupplierName from supplier where shopid = '" & request.cookies("shopid") & "' and length(SupplierName) > 1 and Active = 'YES' order by displayorder, SupplierName"
                      response.write cstmt
                      set crs = con.execute(cstmt)
                      if not crs.eof then
                      	crs.movefirst
                      	do while not crs.eof 
                      		response.write "<option value=""" & crs("SupplierName") & """>" & crs("SupplierName") & "</option>"
                      		crs.movenext
                      	loop
                      else
                      	response.write "<option value='None'>No Suppliers Entered</option>"
                      end if
                      %>
             </select></td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Invoice #:</strong></td>
            <td width="70%" colspan="2">
			<input  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" type="text" name="invoicenumber" size="11" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; value="<%=request("Price")%>" tabindex="9" value=" "></td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Bin/Location</strong></td>
            <td width="70%" colspan="2">
			<input  onfocus="this.select()" type="text" name="bin" size="11" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; width: 163px;"<%=request("bin")%>" tabindex="10" value=" "></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>On Hand:</b></td>
            <td width="70%" colspan="2">
			<input  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" type="text" name="OnHand" size="8" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; value="0" tabindex="11" value="0"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Re-Order Level: </b></td>
            <td width="70%" colspan="2">
			<input type="text" name="ReOrderLevel"  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" size="8" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; tabindex="11" value="0" tabindex="12"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Max On Hand:</b></td>
            <td width="70%" colspan="2">
			<input type="text" name="MaxOnHand"  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC';this.select()" size="8" style="font-family: Arial; font-size: 10pt; text-transform:uppercase; height: 20px;"12" value="0" tabindex="13"></td>
           </tr>
           <tr>
            <td width="30%" align="right"><b>Maintain Stock: </b></td>
            <td width="70%" colspan="2">
			<select size="1" name="MaintainStock" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC'" tabindex="14" >
                <option selected value="Yes">Yes</option>
                <option value="No">No</option>
              </select></td>
           </tr>
           <tr>
            <td width="30%" align="right"><strong>Taxable:
			</strong></td>
            <td width="70%" colspan="2">
			<select size="1" name="tax" onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='#FFFFCC'" tabindex="15" >
                <option selected value="Yes">Yes</option>
                <option value="No">No</option>
              </select></td>
           </tr>
           <tr>
            <td width="100%" colspan="3">
             <p align="center"><input type="button" onclick="checkForm()" name="Abutton3" value="Add to Inventory">&nbsp;<input onclick="location.href='inventory.asp'" name="Button1" type="button" value="Cancel"></td>
           </tr>
          </table></td>
            </tr>
          </table>
         </form>
         </center>
        </div>
       </td>
      </tr>
      <tr>
       <td valign="top" width="100%"></td>
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
document.theform.PartNumber.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>