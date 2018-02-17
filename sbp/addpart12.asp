<!-- #include file=aspscripts/auth.asp --> <!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->

<%
shopid = request.cookies("shopid")
' check for matching parts in partsregistry and partsinventory
stmt = "select partsinventory.partnumber, partsinventory.* from partsinventory where shopid = '" & request.cookies("shopid") & "' and  " _
& "partsinventory.partnumber not in (select partsregistry.partnumber from partsregistry where shopid = '" & request.cookies("shopid") _
& "' and partsinventory.partnumber = partsregistry.partnumber) order by partsinventory.partnumber"
'response.write stmt & "<br>"
set rs = con.execute(stmt)
if not rs.eof then
	do until rs.eof
	
		'add to parts registry
		stmt = "insert into partsregistry (shopid,PartNumber,PartDesc,PartPrice,PartCode,PartCost,PartSupplier,PartCategory,tax,overridematrix)" _
		& " values ('" & shopid & "', '" _
		& replace(rs("partnumber"),"&"," and ") & "', '" _
		& replace(rs("partdesc"),"'","''") & "', " _
		& rs("partprice") & ", '" _
		& rs("partcode") & "', " _
		& rs("partcost") & ", '" _
		& replace(rs("partsupplier"),"'","''") & "', '" _
		& rs("partcategory") & "', '" _
		& rs("tax") & "', '" _
		& rs("overridematrix") & "')"
		'response.write stmt
		con.execute(stmt)
		rs.movenext
		
	loop
end if
		
' get company preference for part search
stmt = "select defaultpartsearch,ignorespaceslashdashpartsearch from company where shopid = '" & request.cookies("shopid") & "'"
set srs = con.execute(stmt)
dps = srs("defaultpartsearch")
ig = srs("ignorespaceslashdashpartsearch")

%>
<html>
<head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<style type="text/css">
<!--


.style1 {
	border-width: 0;
}

.style2 {
	font-weight: bold;
	background-color: #FFFFFF;
}
.style3 {
	background-color: #69A6E3;
}

.style4 {
	font-weight: bold;
	background-color: #FFFFFF;
	text-align: right;
}

.style5 {
	text-align: center;
	background-image: url('https://<%=request.servervariables("SERVER_NAME")%>/sbp/newimages/wipheader.jpg');
}
.style6 {
	color: #FFFFFF;
}
.style7 {
	background-color: #FFFFFF;
}

.auto-style1 {
	text-align: right;
	background-color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
}
.auto-style3 {
	font-family: Arial, Helvetica, sans-serif;
}
.auto-style4 {
	font-weight: bold;
	background-color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
}

-->
</style>

</head>
<script language="JavaScript">

function loadBottom(){

	var randomnumber=Math.floor(Math.random()*11)
	var xmlhttp=new XMLHttpRequest();
	//alert("getting vin explosion")
	var pval = "pfind2.asp?ig=<%=ig%>&sid=<%=request.cookies("shopid")%>&r="+randomnumber+"&roid=<%=request("roid")%>&comid=<%=request("comid")%>"
	<%
	if lcase(dps) <> "both" then
	%>
	sf = document.getElementById("sf").value
	sb = document.getElementById("sb").value
	st = document.getElementById("searchtype").value
	pval = pval + "&sf=" + sf + "&sb=" + sb + "&searchtype=" + st
	<%
	else
	%>
	sb = document.getElementById("sf").value
	st = document.getElementById("searchtype").value
	pval = pval + "&sf=" + sf + "&searchtype=" + st
	<%
	end if
	%>

	xmlhttp.onreadystatechange=function(){
		if (xmlhttp.readyState==4 && xmlhttp.status==200){
			//alert(xmlhttp.responseText)
			document.getElementById("plist").innerHTML = xmlhttp.responseText
		}
	}
	url = pval
	//alert(url)
	xmlhttp.open("GET",url,true);
	xmlhttp.send();

}

	setTimeout("loadBottom()",10)
	
	function addPart(){
	
		location.href='addpart.asp?partnumber='+document.theform.sf.value+'&roid=<%=request("roid")%>&comid=<%=request("comid")%>'

	}
	
	function deletePart(partnumber,sb,sf,comid,roid){

		c = confirm("This will delete this part from inventory.  Are you sure?")
		if (c){
			l = 'deleteinventoryro.asp?comid=<%=request("comid")%>&roid=<%=request("roid")%>&partnumber='+partnumber+'&sb='+sb+'&sf='+sf+'&redir=pfind.asp'
			//alert(l)
			location.href=l
		}

}
	
</script>
<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="132">
    <tr>
      <td width="100%" valign="top" height="89">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="96">
            <tr>
              <td width="110" valign="top" align="left" height="96">
                &nbsp;</td>
              <td valign="top" width="100%" height="96">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style5">
                        <span class="style6"><strong>Enter your Part Number here.&nbsp; 
						If the part is found below, click the part number</strong></span><strong><br class="style6">
						</strong><span class="style6"><strong>to add it to your Repair Order.&nbsp; If not, click on 
						Add New Part to create a new part.</strong></span></td>
                    </tr>
                  </table>
                </div>
  </center>
                <div align="center">
                  <form name="theform" method="get">

                  <table width="60%" cellspacing="0" cellpadding="0" class="style1">
                    <tr>
                      <td width="100%" class="style3">
						<table border="0" width="100%" cellspacing="0" cellpadding="4">
						<%
							if dps <> "Both" then
						%>
                          <tr>
                            <td style="width: 50%" class="auto-style4">
                              <p align="right" class="auto-style3">Search by: </p>
                            </td>
  <center>
                  <center>
                            <td style="width: 50%" class="style7">
							<%
							if dps = "Part Number" then
								nselected = "selected='selected'"
							elseif dps = "Part Description" then
								nselected = ""
							end if
							if dps = "Part Description" then
								dselected = "selected='selected'"
							elseif dps = "Part Number" then
								dselected = ""
							end if
							%>
							<select id="sb" onchange=document.theform.sf.focus() name="sb">
							<option <%=nselected%> value="partnumber">Part Number</option>
							<option <%=dselected%> value="partdesc">Part Description</option>
							</select>
							</td>
                          </tr>
                          <%
                          end if
                          %>
                          <tr>
                            <td style="width: 50%" class="auto-style1">
                              <strong>Search for:                             </strong>                             </td>
                            <td style="width: 50%" class="style7">
                            <input id="searchtype" name="searchtype" value="<%=dps%>" type="hidden" >
							<input id="sf" onkeyup="loadBottom()" onfocus="javascript:this.style.backgroundColor='yellow'" onblur="javascript:this.style.backgroundColor='white'" type="text" name="sf" size="27" style="font-variant: small-caps; font-family: Arial; font-size: 12pt; border-style: solid; border-color: #FFFFFF; width: 218px;" value=""></td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="2" align="center" class="style7">
                            &nbsp;
                            <input type="button" value="Add New Part" onclick="addPart()" name="Abutton1">
		<input type="button" value="Back to RO" onclick="location.href='ro.asp?roid=<%=request("roid")%>'" name="Abutton3">
                            <br>
                            </td>
                          </tr>
                        </table>
                          </center>
                        </center></td>
                    </tr>
                  </table><input id="roid" type="hidden" name="roid" value="<%=request("roid")%>">
                  <input id="comid" type=hidden name=comid value="<%=request("comid")%>">
                  </form>
                </div>

            </td>
            </tr>
          </table>
        </div>
    </td>
    </tr>
  </table>
</div>
<div id="plist" style="width:100%;height:500px;border:0px"></plist>
</body>
<script language="javascript">
document.theform.sf.focus()

function getDocHeight() {
    var D = document;
    return Math.max(
        Math.max(D.body.scrollHeight, D.documentElement.scrollHeight),
        Math.max(D.body.offsetHeight, D.documentElement.offsetHeight),
        Math.max(D.body.clientHeight, D.documentElement.clientHeight)
    );
}

h = getDocHeight()
nh = h-200
//alert(nh)
//document.getElementById("results").style.height = (nh)
document.getElementById("plist").style.height = nh

</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>