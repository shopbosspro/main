<!-- #include file=aspscripts/auth.asp --> <!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->

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

' check for matching parts in partsregistry and partsinventory
stmt = "select partsinventory.partnumber, partsinventory.* from partsinventory where shopid = '" & shopid & "' and  " _
& "partsinventory.partnumber not in (select " & preg & ".partnumber from " & preg & " where shopid = '" & shopid _
& "' and partsinventory.partnumber = " & preg & ".partnumber) order by partsinventory.partnumber"
'response.write stmt & "<br>"
set rs = con.execute(stmt)
if not rs.eof then
	do until rs.eof
	
		'add to parts registry
		stmt = "insert into " & preg & " (shopid,PartNumber,PartDesc,PartPrice,PartCode,PartCost,PartSupplier,PartCategory,tax,overridematrix)"
		stmt = stmt & " values ('" & shopid & "', '"
		stmt = stmt & replace(replace(rs("partnumber"),"&"," and "),"'","''") & "', '"
		stmt = stmt & replace(rs("partdesc"),"'","''") & "', "
		stmt = stmt & rs("partprice") & ", '"
		stmt = stmt & rs("partcode") & "', "
		stmt = stmt & rs("partcost") & ", '"
		stmt = stmt & replace(rs("partsupplier"),"'","''") & "', '"
		stmt = stmt & rs("partcategory") & "', '"
		stmt = stmt & rs("tax") & "', '"
		stmt = stmt & rs("overridematrix") & "')"
		'response.write stmt
		con.execute(stmt)
		rs.movenext
		
	loop
end if
		
' get company preference for part search
stmt = "select defaultpartsearch,ignorespaceslashdashpartsearch,usepartsregistry from company where shopid = '" & request.cookies("shopid") & "'"
set srs = con.execute(stmt)
dps = srs("defaultpartsearch")
upr = srs("usepartsregistry")
ig = srs("ignorespaceslashdashpartsearch")


shopid = request.cookies("shopid")
stmt = "select * from joinedshops where shopid = '" & shopid & "' or joinedshopid = '" & shopid & "'"
set rs = con.execute(stmt)
if not rs.eof then
	do until rs.eof
		jslist = jslist & "'" & rs("shopid") & "',"
		rs.movenext
	loop
else
	jslist = "'" & shopid & "'"
end if
if instr(jslist,",") > 0 then
	jslist = server.urlencode(left(jslist,len(jslist)-1))
end if



%>
<html>
<head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<script src="jquery/jquery-1.10.2.min.js"></script>
<script src="jquery/jquery.cookie.js"></script>
<script src="jquery/jquery.typewatch.js"></script>
<script src="javascripts/bootstrap.min.js"></script>
<link href="css/bootstrap.min.css" rel="stylesheet">
<style type="text/css">
<!--

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }
button{
	color:maroon;
	font-size:small;
	font-weight:bold;
	cursor:hand
}
.menubutton{
	width:88px;
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
	background-color: #FFFFFF;
	text-align: center;
}

-->
</style>

</head>
<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">

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
                        <span class="style6"><strong>Enter your Part Number 
						here.&nbsp; If the part is found below, click the part number</strong></span><strong><br class="style6">
						</strong><span class="style6"><strong>to add it to your 
						Repair Order.&nbsp; If not, click on Add New Part to create a 
						new part.</strong></span></td>
                    </tr>
                  </table>
                </div>
                <div align="center">
                  <form name="theform" onsubmit="return false;" method="get">

                  <table width="60%" cellspacing="0" cellpadding="0" class="style1">
                    <tr>
                      <td width="100%" >
                      <%
                      stmt = "select vehinfo,customer,vehengine,vehtrans,cyl from repairorders where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
                      set rors = con.execute(stmt)
                      response.write "<div style='text-align:center;font-size:14px;font-weight:bold;padding:15px;'>" & rors("customer") & " - " & rors("vehinfo") & " " & rors("vehengine") & " " & rors("vehtrans") & "</div>"
                      set rors = nothing
                      %>
						<table class="table table-condensed table-striped">
							<tr>
                            <td style="width: 50%" class="style4">
                              Search for:                             </td>
                            <td style="width: 50%" class="style7">
                            <input id="searchtype" name="searchtype" value="<%=dps%>" type="hidden" >
							<input id="sf" 
							onfocus="javascript:this.style.backgroundColor='yellow'" 
							onblur="javascript:this.style.backgroundColor='white'" 
							type="text" name="sf" style="padding:5px;border:1px silver solid;border-radius:5px;width:200px;text-transform:uppercase" 
							value=""></td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="2" align="center" class="style7">
                            &nbsp;
                            <input type="button" class="btn btn-primary btn-md" value="Add New Part" onclick="addPart()" name="Abutton1">
		<input type="button" value="Back to RO" class="btn btn-info btn-md" onclick="location.href='ro.asp?roid=<%=request("roid")%>'" name="Abutton3">
                            <br>
                            </td>
                          </tr>
                        </table>
                       </td>
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
<img src="newimages/loaderbig.gif" style="display:none;width:75px;height:75px;position:absolute;left:48%;top:300px;" id="spinner">
<div id="results"></div>
</body>
<script language="javascript">

$(document).ready(function(){

	$('#sf').focus()
	var twoptions = {
	    callback: function(){
	    	$('#spinner').toggle()
			sf = $('#sf').val()
	    	ds = "upr=<%=upr%>&ig=<%=ig%>&sid=<%=jslist%>&roid=<%=request("roid")%>&comid=<%=request("comid")%>&sf=" + sf
	    	$.ajax({
	    		data: ds,
	    		url: "pfind.asp",
	    		success: function(r){
	    			$('#results').html(r)
	    			$('#spinner').toggle()
	    		}
	    	});
	    },
	    wait: 150,
	    highlight: true,
	    allowSubmit: false,
	    captureLength: 2
	}
	
	$("#sf").typeWatch( twoptions );
	
	sf = $('#sf').val()
	$('#spinner').toggle()
	ds = "upr=<%=upr%>&ig=<%=ig%>&sid=<%=jslist%>&roid=<%=request("roid")%>&comid=<%=request("comid")%>&sf=" + sf
	$.ajax({
		data: ds,
		url: "pfind.asp",
		success: function(r){
			$('#results').html(r)
			$('#spinner').toggle()
		}
	});

});

function addPart(){

	location.href='addpart.asp?partnumber='+document.theform.sf.value+'&roid=<%=request("roid")%>&comid=<%=request("comid")%>'

}

function deletePart(partnumber,sb,sf,comid,roid){

	c = confirm("This will delete this part from inventory.  Are you sure?")
	if(c){
		$.ajax({
		
			data: "partnumber="+partnumber+"&shopid=<%=request.cookies("shopid")%>",
			url: "deleteinventoryro.asp",
			success: function(r){
				sf = $('#sf').val()
				$('#spinner').toggle()
				ds = "upr=<%=upr%>&ig=<%=ig%>&sid=<%=jslist%>&roid=<%=request("roid")%>&comid=<%=request("comid")%>&sf=" + sf
				$.ajax({
					data: ds,
					url: "pfind.asp",
					success: function(r){
						$('#results').html(r)
						$('#spinner').toggle()
					}
				});
			}
		
		});
	}


}

</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>