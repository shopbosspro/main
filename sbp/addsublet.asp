<!-- #include file=aspscripts/auth.asp --> 
 <!-- #include file=aspscripts/conn.asp --> 


<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" src="javascripts/killenter.js"></script>
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

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style2 {
	font-weight: bold;
}

.style3 {
	border: 1px solid #000000;
}

.style4 {
	font-size: medium;
	color: #FF0000;
}

.style5 {
	color: #FFFFFF;
}
.style6 {
	font-weight: bold;
		background-image: url('newimages/wipheader.jpg');
}
#hider{
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

-->
</style>
<script type="text/javascript">
function checkForm(){
	d = document.theform
	if(d.Sublet.value.length == 0){
		alert("Sublet is required")
		return false;
	}
	if(d.Supplier.value.length == 0){
		alert("Supplier is required")
		return false;
	}
	if(d.Price.value.length == 0){
		alert("Price is required")
		return false;
	}
	if(d.Cost.value.length == 0){
		alert("Cost is required")
		return false;
	}


	document.theform.submit()
}

function createNewPO(){
	
	r = Math.random()
	document.getElementById("newpo").src = "addsubletnewpo.asp?roid=<%=request("roid")%>&r=" + r
	document.getElementById("newpo").style.display = 'block'
	document.getElementById("hider").style.display = 'block'

}
</script>

</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
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
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style6">
                        <p align="center"><span class="style5">Add your sublet item here, then click
                        on Add Sublet.</span><br class="style5">
						<span class="style5">All fields are required</span></td>
                    </tr>
                  </table>
                </div>
                <div align="center">
                  <center>
                  <form method="post" action="addsubletnow.asp" name="theform">
                  <table width="70%" cellspacing="0" cellpadding="2" class="style3">
                    <tr>
                      <td width="30%" align="right" class="style2">Sublet Item</td>
                      <td width="70%"><input type="text" name="Sublet" size="63" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("Sublet")%>"><strong><span class="style4">*</span></strong></td>
                    </tr>
                    <tr>
                      <td width="30%" align="right" class="style2">Selling Price</td>
                      <td width="70%"><input type="text" name="Price" size="11" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("Price")%>"><strong><span class="style4">*</span></strong></td>
                    </tr>
                    <tr>
                      <td width="30%" align="right" class="style2">Supplier</td>
                      <td width="70%"><input type="text" name="Supplier" size="11" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("Supplier")%>"><strong><span class="style4">*</span></strong></td>
                    </tr>
		           <tr>
		            <td width="30%" align="right" class="style2"><span class="style7">
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
                      <td width="30%" align="right" class="style2">Invoice
                        #</td>
                      <td width="70%"><input value="<%=request("Invoice")%>" type="text" name="Invoice" size="11" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px"></td>
                    </tr>
                    <tr>
                      <td width="30%" align="right" class="style2">Shop Cost</td>
                      <td width="70%"><input value="<%=request("Cost")%>" type="text" name="Cost" size="11" style="font-family: Arial; font-size: 12pt; font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px"><strong><span class="style4">*</span></strong></td>
                    </tr>
                    <tr>
                      <td width="100%" colspan="2">
                        <p align="center"><input value="Add Sublet" onclick="checkForm()" type="button" class="obutton" name="Abutton1">
						<input type="button" value="Cancel" class="obutton" onclick="location.href='ro.asp?roid=<%=request("roid")%>'" name="Abutton6">
				

                      </td>
                    </tr>
                  </table><input type="hidden" name="roid" value="<%=request("roid")%>">
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
document.theform.Sublet.focus()
</script>
<p>
&nbsp;</p>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>