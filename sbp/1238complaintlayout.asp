<!-- #include file=aspscripts/conn.asp -->
<head>
<style type="text/css">
body{
margin:0;
}
.style2222 {
	text-align: left;
	background-image:url('newimages/pageheader.jpg');
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 14px;
	color: white;
	}

}
.style2223 {
	font-weight: bold;
}
.style2224 {
	text-align: right;
}
.style2225 {
	text-align: left;
	z-index:999;
}
.style2226 {
	text-decoration: underline;
}
.style2227 {
	color: #FFFFFF;
}
</style>
<script type="text/javascript">
function findPos(obj,comid) {

  if (document.getElementById("drop"+comid).style.display == "inline-block"){
  	document.getElementById("drop"+comid).style.display = "none"
  	return
  }
  
  var curleft = curtop = 0, scr = obj, fixed = false;
  while ((scr = scr.parentNode) && scr != document.body) {
    curleft -= scr.scrollLeft || 0;
    curtop -= scr.scrollTop || 0;
    if (getStyle(scr, "position") == "fixed") fixed = true;
  }
  if (fixed && !window.opera) {
    var scrDist = scrollDist();
    curleft += scrDist[0];
    curtop += scrDist[1];
  }
  do {
    curleft += obj.offsetLeft;
    curtop += obj.offsetTop;
  } while (obj = obj.offsetParent);
  //return [curleft, curtop];
  //document.getElementById("showcoord").value=curleft+"*"+curtop
  
  // now show the menu
  if(navigator.appName == "Microsoft Internet Explorer"){
  	  newcurleft = curleft-38
  	  newcurtop = curtop + 26
  	  //alert(newcurleft+":"+newcurtop)
	  document.getElementById("drop"+comid).style.left = newcurleft+"px"
	  document.getElementById("drop"+comid).style.top = newcurtop+"px"
	  document.getElementById("drop"+comid).style.display = 'inline-block'
  }else{
  	  newcurleft = curleft+6
  	  newcurtop = curtop+17
	  document.getElementById("drop"+comid).style.left = newcurleft+"px"
	  document.getElementById("drop"+comid).style.top = newcurtop+"px"
	  document.getElementById("drop"+comid).style.display = 'inline-block'
  }

}

function scrollDist() {
  var html = document.getElementsByTagName('html')[0];
  if (html.scrollTop && document.documentElement.scrollTop) {
    return [html.scrollLeft, html.scrollTop];
  } else if (html.scrollTop || document.documentElement.scrollTop) {
    return [
      html.scrollLeft + document.documentElement.scrollLeft,
      html.scrollTop + document.documentElement.scrollTop
    ];
  } else if (document.body.scrollTop)
    return [document.body.scrollLeft, document.body.scrollTop];
  return [0, 0];
}

function getStyle(obj, styleProp) {
  if (obj.currentStyle) {
    var y = obj.currentStyle[styleProp];
  } else if (window.getComputedStyle)
    var y = window.getComputedStyle(obj, null)[styleProp];
  return y;
}



</script>
</head>
<input type="hidden" id="showcoord" value="">
<div >

<div class="style2222">
	<table cellpadding="0" cellspacing="0" style="width: 100%">
		<tr>
			<td style="width: 50%">
			<input id="issuebutton1" style="float:right; height: 33px; width: 200px;border-radius:5px; border:1px black solid; -webkit-box-shadow: 5px 5px 5px 0px rgba(0,0,0,0.75);
-moz-box-shadow: 5px 5px 5px 0px rgba(0,0,0,0.75);
box-shadow: 5px 5px 5px 0px rgba(0,0,0,0.75);background-color:#006600;color:white;cursor:pointer" type="button" value="Change Vehicle Issue Order" onclick="changeVIOrder()">

	<img alt="" height="30" src="newimages/vintage-car-icon.png" width="45"></td>
<%
if len(request("shopid")) > 0 then
	response.cookies("shopid") = request("shopid")
end if
shopid = request.cookies("shopid")
stmt = "select lineitemcomplete, showpartsbuttons, alldatausername, alldatapassword from company where shopid = '" & shopid & "'"
set myrs = con.execute(stmt)
lineitemcomplete = lcase(myrs("lineitemcomplete"))
if len(myrs("alldatausername")) > 0 and len(myrs("alldatapassword")) > 0 then
	egbutton = "yes"
else
	egbutton = "no"
end if

if myrs("showpartsbuttons") = "yes" then
%>
			<td style="width: 50%"><span class="style2227"><strong>Parts Ordering </strong></span>&nbsp;<input onclick="showParts('http://weblink.carquest.com')" name="Button1" type="button" value="CarQuest">
			<input onclick="showParts('https://www.autozonepro.com/')" name="Button2" type="button" value="AutoZone">
			<input onclick="showParts('https://www.napaonline.com/YourAccount/login.aspx?ref=header')" name="Button3" type="button" value="NAPA">
			<input onclick="showParts('https://fco.firstcallonline.com/FirstCallOnline/public.html')" name="Button4" type="button" value="O'Reilly">
			<%
			if shopid = "1055" then
			%>
			<input onclick="showParts('https://www.autopartintl.com/')" name="Button4" type="button" value="Auto Part Intl">
			<%
			end if
			%>
			</td>
<%
end if
%>
		</tr>
	</table>
	
	Vehicle 
	Issues (Incomplete parts/labor will be shown in <span style="color:#FFFF99">yellow</span>.  Click the <img alt="" height="14" src="newimages/greencheckmark.jpg" width="17"> to mark an item complete)</div>

<%
'*** test login and password for estimating guide button


function getcolor()
  ccstmt = "select partid from parts where shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  set ccrs = con.execute(ccstmt)
  if not ccrs.eof then
    ccolor = "#FFFFFF"
    exit function
  end if
  set ccrs = nothing
  'check for labor
  llstmt = "select laborid from labor where shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  set llrs = con.execute(llstmt)
  if not llrs.eof then
    ccolor = "#FFFFFF"
    exit function
  end if
  set llrs = nothing
  'check for sublet
  ssstmt = "select subletid from sublet where deleted = 'no' and shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  set ssrs = con.execute(ssstmt)
  if not ssrs.eof then
    ccolor = "#FFFFFF"
    exit function
  end if
  set ssrs = nothing
end function


'get complaints
roid = request("roid")
stmt = "select complaintid from complaints where cstatus = 'no' and shopid = '" & shopid & "' and roid = " & roid
set comrs = con.execute(stmt)
comc = 0
if not comrs.eof then
	do until comrs.eof
		comc = comc + 1
		comrs.movenext
	loop
end if

stmt = "select id from complaints where cstatus = 'no' and shopid = '" & shopid & "' and roid = " & roid
set comrs = con.execute(stmt)
if not comrs.eof then
	comc = 0
	do until comrs.eof
		comc = comc + 1
		comrs.movenext
	loop
else
	comc = 0
end if


stmt = "select * from complaints where cstatus = 'no' and shopid = '" & shopid & "' and roid = " & roid & " order by displayorder"
set comrs = con.execute(stmt)
cntr = 1
if not comrs.eof then
  comrs.movefirst
  if cntr = 1 then
  tcolor = cntr mod 2
  if tcolor = 0 then
  	mbgc = "'#0066CC'"
  else
  	mbgc = "'#0066CC'"
  end if
  end if
  while not comrs.eof
  comid = comrs("complaintid")
  comrootid = comrs("id")
  ccolor = "#FFFF00"
  ad = comrs("acceptdecline")
  if ad = "Declined" then
  	itemcolor = ";background-color:black;color:white'"
  elseif ad = "Pending: Parts Ordered" then
  	itemcolor = ";background-color:#FF9;color:black'"
  elseif ad = "Approved" then
  	itemcolor = ";background-color:maroon;color:white'"
  elseif ad = "Parts Ordered" then
  	itemcolor = ";background-color:red;color:white'"
  elseif ad = "Assembly" then
  	itemcolor = ";background-color:green;color:white'"
  elseif ad = "Job Complete" then
  	itemcolor = ";background-color:navy;color:white'"
  else
  	itemcolor = ""
  end if
  	
  'getcolor()
%><table border=1 cellspacing=0 cellpadding=0 width=100% bordercolor="#000000"><tr><td>
<table background="newimages/wipheader.jpg" bgcolor=<%=mbgc%> width="100%" cellspacing="0" bordercolor="#99CCFF" cellpadding="2" style="border-collapse: collapse">
  <tr>
    <td style="width: 90%" class="style2223"><font size="2"><font color="#FFFFFF">
    <a name="<%=comrs("displayorder")%>"></a>#
    <%=comrs("displayorder")%>
    . &nbsp;<font color="<%=ccolor%>"><%=ucase(comrs("complaint"))%>
    <%
    if len(comrs("acceptdecline")) = 0 or comrs("acceptdecline") = "Accept" then
    %>
    &nbsp;<%
    elseif comrs("acceptdecline") = "Decline" then
    %>
    <%
	end if
	%>
	 </font><br>
      Tech Story: <%=ucase(comrs("techreport"))%> <span onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';saveAll('techreport.asp?comrootid=<%=comrootid%>&comid=<%=comid%>&')" style="color:yellow;cursor:pointer">[more]</span></font>
      
    
    <br>

    </td>
    <td style="width: 10%" class="style2224">
    <%
	if cntr > 1 and cntr <= comc then
	%>
	<img style="cursor:pointer" onclick="setCookie('<%=comrs("displayorder")%>');changePos('<%=comrs("id")%>','<%=comrs("displayorder")%>','up','<%=comrs("roid")%>')" alt="" height="30" src="newimages/uparrow.png" width="30">
	<%
	end if
	if cntr >= 1 and cntr < comc then
	%>
	<img style="cursor:pointer" onclick="setCookie('<%=comrs("displayorder")%>');changePos('<%=comrs("id")%>','<%=comrs("displayorder")%>','down','<%=comrs("roid")%>')"  alt="" height="30" src="newimages/downarrow.png" width="30">
	<%
	end if
	%>
	</td>
  </tr>
  </table>
  <!-- parts table here -->
  <table width=100% bgcolor=white cellpadding=2 cellspacing=1><tr>
	<td valign="top">
  <%
  'get parts
  cstmt = "select pstatus, `shopid`,`PartID`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax`,`overridematrix`,`posted` from parts where shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  addbut = "no"
  set crs = con.execute(cstmt)

  if egbutton = "yes" then
	  response.write "<input type='button' style='width:170px;font-weight:bold' class='abuts' onclick='showGuide(""" & roid & """,""" & comid & """,""y"")' value='Parts & Labor Guide'>"
  end if
  response.write "<input style='width:65px;' type='button' value='Parts' class='abuts' onclick=setCookie('" & comrs("displayorder") & "');document.body.onunload='',saveAll('addpart1.asp?comid=" & comid & "&','part','','','')>"
  response.write "<input style='width:65px;' type='button' value='Labor' class='abuts' onclick=setCookie('" & comrs("displayorder") & "');document.body.onunload='';saveAll('addlabor.asp?comid=" & comid & "&','labor','','','')>"
  response.write "<input style='width:65px;' type='button' value='Sublet' class='abuts' onclick=setCookie('" & comrs("displayorder") & "');document.body.onunload='';saveAll('addsublet.asp?comid=" & comid & "&','sublet','','','')>"
  response.write "<input style='width:85px;' type='button' value='Canned Job' style='width:100px;' class='abuts' onclick=setCookie('" & comrs("displayorder") & "');document.body.onunload='';saveAll('addcannedjobtoro.asp?comid=" & comid & "&','canned','','','')>"
  response.write "<input style='width:85px;' class='abuts' onclick=setCookie('" & comrs("displayorder") & "');document.body.onunload='';saveAll('techreport.asp?comrootid=" & comrootid & "&comid=" & comid &"&','','','','') style='font-size:12px;height:24px;width:100px;' type='button' value='Tech Story' name='ts'>"
  %>
  &nbsp;<input onclick="setCookie('<%=comrs("displayorder")%>');findPos(this,'<%=comrs("complaintid")%>')" id="dropbutton<%=comrs("complaintid")%>" type='button' value='Status:<%=comrs("acceptdecline")%>' style='cursor:pointer;width:210px; font-weight:bold<%=itemcolor%>' class='abuts' onclick=document.body.onunload='';saveAll('addcannedjobtoro.asp?comid=" & comid & "&','canned','','','')>
  <%
  select case comrs("acceptdecline")
  	case "Approved"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
  <%
  	case "Pending"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
    <li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
  <%
  	case "Pending: Parts Ordered"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>

  <%
  	case "Declined"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
      <li onclick="document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>      
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
  <%
  	case "Parts Ordered"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
      <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>      
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
  <%
  	case "Assembly"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
      <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>      
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
   	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
  <%
  	case "Job Complete"
  %>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
      <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>      
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>
	<%
	case else
	%>
  <ul id="drop<%=comrs("complaintid")%>" style="list-style-type:none;display:none;position:absolute;width:130px;background-color:silver;z-index:1001;font-weight:bold; left: 0px; top: 0px; padding:5px;border:thin gray solid" id="test">
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending</li>
    <li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Pending: Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Pending: Parts Ordered</li>    
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Approved';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Approved</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');declineItem('<%=comrs("complaintid")%>')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Declined</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Parts Ordered';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Parts Ordered</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Assembly';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Assembly</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';document.theform.action='1238savero.asp?comid=<%=comrs("complaintid")%>&comstat=Job Complete';saveAll('ro.asp?a=n','','','','','')" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Job Complete</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.body.onunload='';deleteVehIssue('<%=comrs("complaintid")%>');" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer">Delete Issue</li>
  	<li onclick="setCookie('<%=comrs("displayorder")%>');document.getElementById('drop<%=comrs("complaintid")%>').style.display='none'" onmouseout="this.style.backgroundColor='silver';this.style.color='black'" onmouseover="this.style.backgroundColor='green';this.style.color='white'" style="padding:5px;cursor:pointer;font-size:x-small">Close This</li>
  </ul>

  <%
  end select
  response.write "<span style='font-size:small;color:white;font-weight:normal'>"

  
  if crs.eof then
  	response.write "</td></tr></table>"
  else
  	crs.movefirst
  	comttlprts = 0
  	if addbut = "yes" then
  		response.write "<button class='abuts' onclick='orderPart(" & roid & ")'>Order</button>"
  	end if
  	response.write "</td></tr></table>"
  %>
  <table border=1 cellspacing=0 width=100% bordercolor="#99CCFF" style="border-collapse: collapse;" cellpadding="2" bgcolor="#FFFFFF">
  <%
  while not crs.eof
  	if lineitemcomplete = "yes" then
	  	if lcase(crs("pstatus")) <> "done" then
	  		backcolor = "#FF9"
	  	else
	  		backcolor = "#CFC"
	  	end if
	else
		backcolor = "white"
	end if
  
  %>
  <tr>
    <td style="background-color:<%=backcolor%>" width="19%"><font size="1">
	<img style="cursor:pointer" onclick="setCookie('<%=comrs("displayorder")%>');completeItem('part','<%=crs("partid")%>','<%=crs("pstatus")%>')" alt="" height="14" src="newimages/greencheckmark.gif" width="17"><b>(P)</b> <%=trim(crs("partnumber"))%>&nbsp;</font></td>
    <td width="42%"><font size="1"><%=crs("partdesc")%>&nbsp;</font></td>
    <td width="17%" align="right">
      <p align="center"><font size="1"><%=crs("quantity")%>  @ <%=formatcurrency(crs("partprice"))%>
      <%
      if ucase(crs("tax")) = "YES" then 
      	response.write "(T)"
      else
      	response.write "(NT)"
      end if
      %>
      </font></p>
    </td>
    <td width="8%" align="right"><font color=red size="1">
      <%=crs("discount")%>%</td>
    <td width="9%" align="right"><font size="1"><%=formatcurrency(crs("linettlprice"))%></font>&nbsp;</td>
    <td width="5%" align="center"><font size="1"><a href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';saveAll('editpart.asp?','part','','<%=crs("PartID")%>','')"><img border=0 src=newimages/edit2.png width=16 height=17></a></font><a onmouseover="javascript:status='Delete Part';return true" onmouseout="javascript:status='Shop Boss Pro';return true" href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';conDel('delroitem.asp?','part','','<%=crs("PartID")%>','')"><img border="0" src="newimages/delete.png" width="16" height="17"></a>
	</td>
  </tr>
  <%
  	comttlprts = comttlprts + crs("linettlprice")
  	'response.write crs("tax")
  	if ucase(crs("tax")) = "YES" then
  		taxparts = taxparts + crs("linettlprice")
  	else
  		notaxparts = notaxparts + crs("linettlprice")
  	end if
  	'response.write taxparts
    crs.movenext
   wend
   %>
   </table>
   <%   
  end if
  'get labor
  cstmt = "select lstatus,`shopid`,`LaborID`,`ROID`,`HourlyRate`,`LaborHours`,`Labor`,`Tech`,`LineTotal`,`LaborOp`,`complaintid`,`techrate` from labor where shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  set crs = con.execute(cstmt)
  if crs.eof then
  	'response.write "&nbsp;&nbsp;&nbsp;<b><a href=""javascript:document.body.onunload='';saveAll('addlabor.asp?comid=" & comid & "&','labor','','','')""><font color=#0066cc>Labor</a>"
  else
  	crs.movefirst
  	comttllbr = 0
  	'response.write "&nbsp;&nbsp;&nbsp;<b><a href=""javascript:document.body.onunload='';saveAll('addlabor.asp?comid=" & comid & "&','labor','','','')""><font color=#0066cc>Labor</a>"
  %>
 <!-- parts ends, labor begins -->
  <table border=1 cellspacing=0 width=100% cellpadding="2" bordercolor="#99CCFF" style="border-collapse: collapse;" bgcolor="#FFFFFF">
  <%
    while not crs.eof
    tlc = tlc + crs("techrate")
    mypos = instr(crs("Tech"), ",")
        if mypos = 0 then
          'get employee
          posstmt = "select EmployeeLast from employees where shopid = '" & shopid & "' and EmployeeID = " & cint(crs("Tech"))
	      'response.write posstmt
	      set posrs = con.execute(posstmt)
	      posrs.movefirst
	      thetech = posrs("EmployeeLast")
	      if msplit = "yes" then thetech = thetech & "<BR>SPLIT"
	      set posrs = nothing
        else
	      thetech = left(crs("Tech"), mypos-1)
        end if 
        
        ' check for posted hours by technician
        stmt = "select `id`,`startdatetime`,`enddatetime`,`roid`,`complaintid`,`laborid`,`shopid`,`tech` from labortimeclock where shopid = '" & shopid & "' and roid = " & request("roid") & " and complaintid = " _
        & crs("complaintid") & " and laborid = " & crs("laborid")
        set ltrs = con.execute(stmt)
        ttlclockdiff = 0
        if not ltrs.eof then
        	do until ltrs.eof
        		starttime = ltrs("startdatetime")
        		endtime = ltrs("enddatetime")
        		mindiff = datediff("n",starttime,endtime)
        		ttlclockdiff = ttlclockdiff + mindiff
        		ltrs.movenext
        	loop
        end if
        if ttlclockdiff > 0 then
        	if ttlclockdiff >= 60 then
        		ttlclockdiffdisplay = ttlclockdiff / 60
        		tclockdisplay = "Clock Hrs:" & round(ttlclockdiffdisplay,2)
        	else
        		tclockdisplay = "Clock Mins:" & ttlclockdiff
        	end if
        else
        	tclockdisplay = "Clock Mins:0"
        end if
        if isnull(tclockdisplay) then
        	tclockdisplay= "Clock Mins:0"
        end if
	  	if lineitemcomplete = "yes" then
		  	if lcase(crs("lstatus")) <> "done" then
		  		backcolor = "#FF9"
		  	else
		  		backcolor = "#CFC"
		  	end if
		else
			backcolor = "white"
		end if

  
  %>
  <tr>
    <td style="background-color:<%=backcolor%>" valign="top" width="19%"><font size="1"><b><img style="cursor:pointer" onclick="setCookie('<%=comrs("displayorder")%>');completeItem('labor','<%=crs("laborid")%>','<%=crs("lstatus")%>')" alt="" height="14" src="newimages/greencheckmark.gif" width="17">(L)</b> <%=thetech%>&nbsp;
    <%
    if ttlclockdiff > 0 then
    %>
    <div id="clocklist1-<%=crs("laborid")%>" onclick="showClockList('<%=crs("laborid")%>',this)" style="font-size:12px;height:16px;margin-top:-8px;text-align:right;cursor:pointer;text-decoration:underline;color:#016CCB"><%=tclockdisplay%></div></font>
    <div id="clocklist2-<%=crs("laborid")%>" style="background-color:white;position:relative;width:600px;height:300px;display:block;color:white;border:medium navy outset;overflow:auto;display:none;z-index:999">
    <%
        stmt = "select `id`,`startdatetime`,`enddatetime`,`roid`,`complaintid`,`laborid`,`shopid`,`tech` from labortimeclock where shopid = '" & shopid & "' and roid = " & request("roid") & " and complaintid = " _
        & crs("complaintid") & " and laborid = " & crs("laborid")
    	set ltrs = con.execute(stmt)
    	if not ltrs.eof then
    		
    %>
    	<img onclick="closeClockList('<%=crs("laborid")%>')" alt="" src="newimages/close-icon.png" style="float:right;cursor:pointer">
    	<table style="width: 100%">
			<tr>
				<td class="style2226">Tech&nbsp;</td>
				<td class="style2226">Start Date/Time&nbsp;</td>
				<td class="style2226">End Date/Time&nbsp;</td>
				<td class="style2226">Hours&nbsp;</td>
			</tr>
	<%
			tmindiff = 0
			tmins = 0
			mindiff = 0
			do until ltrs.eof
        		starttime = ltrs("startdatetime")
        		endtime = ltrs("enddatetime")
        		mindiff = datediff("n",starttime,endtime)
        		if mindiff >= 60 then
        			tclockdisplay = "Hours:" & round(mindiff / 60,2)
        		else
        			tclockdisplay = "Mins:" & mindiff
        		end if
        		tmindiff = tmindiff + mindiff
        		tmins = tmins + mindiff

	%>
			<tr>
				<td><%=ltrs("tech")%>&nbsp;</td>
				<td><%=ltrs("startdatetime")%>&nbsp;</td>
				<td><%=ltrs("enddatetime")%>&nbsp;</td>
				<td><%=tclockdisplay%>&nbsp;</td>
			</tr>
	<%
				ltrs.movenext
			loop
			if tmindiff > 0 then
				if tmindiff >= 60 then
					tmindiffdisplay = round(tmindiff / 60,2)
				else
					tmindiffdisplay = tmindiff
				end if
			else
				tmindiff = 0
			end if
			if tmindiff > 0 then
				tmindiffcalc = round(tmindiff / 60,2)
				jobeff = round((crs("laborhours") / tmindiffcalc) * 100,2)
			else
				jobeff = 0
			end if
	%>
			<tr>
				<td><b>Job Efficiency</b>&nbsp;</td>
				<td><b><%=jobeff%>%</b>&nbsp;</td>
				<td><b>Total Hours</b>&nbsp;</td>
				<td><b><%=tmindiffcalc%></b>&nbsp;</td>
			</tr>
		</table>
    <%
    	end if
    %>
    </div>
    <%
   	else
		tmins = 0
		mindiff = 0
		tmindiff = 0
    end if
    if tmins < 1 then
    	backcolor = "style='background-color:yellow;font-weight:bold'"
    else
    	backcolor = "style='background-color:white'"
    end if
    %>
    </td>
    <td width="42%" valign="top" align="right">
      <p align="left"><font size="1"><%=crs("labor")%></font></td>
    <td width="17%" valign="top" align="right">
      <p align="center"><font size="1"><%=crs("laborhours")%>   @ <%=formatcurrency(crs("hourlyrate"))%></font></p>
    </td>
    <td <%=backcolor%> width="8%" valign="top" class="style2225">
      <div id="clock<%=crs("laborid")%>" onclick="timeClock('<%=crs("laborid")%>','<%=crs("complaintid")%>','<%=crs("tech")%>')" style="cursor:pointer;position:relative;z-index:999"><img alt="" src="newimages/clock.png">Clock In</div>

    </td>
    <td width="9%" valign="top" align="right"><font size="1"><%=formatcurrency(crs("linetotal"))%></font>&nbsp;</td>
    <td width="5%" valign="top" align="center"><font size="1"><a href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';saveAll('editlabor.asp?','labor','','','<%=crs("LaborID")%>')"><img border=0 src=newimages/edit2.png width=16 height=17></a></font><a onmouseover="javascript:status='Delete Labor Item';return true" onmouseout="javascript:status='Shop Boss Pro';return true" href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';conDel('delroitem.asp?','labor','','','<%=crs("LaborID")%>')"><img border="0" src="newimages/delete.png" width="16" height="17"></a>
	</td>
  </tr>
  <%
  	ttllbr = ttllbr + crs("linetotal")
  	comttllbr = comttllbr + crs("linetotal")
    crs.movenext
   wend
   %>
   </table>
   <%
  end if
  'get sublet
  cstmt = "select `shopid`,`SubLetID`,`ROID`,`SubletDesc`,`SubletPrice`,`SubletCost`,`SubletInvoiceNo`,`SubletSupplier`,`complaintid` from sublet where deleted = 'no' and shopid = '" & shopid & "' and roid = " & roid & " and complaintid = " & comid
  set crs = con.execute(cstmt)
  if crs.eof then
  	'response.write "&nbsp;&nbsp;&nbsp;<b><a href=""javascript:document.body.onunload='';saveAll('addsublet.asp?comid=" & comid & "&','sublet','','','')""><font color=#0066cc>Sublet</a>"
  else
  	crs.movefirst
  	comttlsub = 0
  	'response.write "&nbsp;&nbsp;&nbsp;<b><a href=""javascript:document.body.onunload='';saveAll('addsublet.asp?comid=" & comid & "&','sublet','','','')""><font color=#0066cc>Sublet</a>"
  %>
  <!-- labor ends, sublet begins -->
  <table border=1 cellspacing=0 width=100% cellpadding="2" bordercolor="#99CCFF" style="border-collapse: collapse" bgcolor="#FFFFFF">
  <%while not crs.eof%>
  <tr>
    <td width="19%"><font size="1"><b>(S)</b> <%=crs("subletinvoiceno")%>&nbsp;</font></td>
    <td width="42%"><font size="1"><%=crs("subletdesc")%>&nbsp;</font></td>
    <td width="17%">
    <p align="center"><font size="1">&nbsp;</font></td>
    <td width="8%">
    &nbsp;</td>
    <td width="9%">
      <p align="right"><font size="1"><%=formatcurrency(crs("subletprice"))%></font></td>
    <td width="5%" align="center"><font size="1"><a href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';saveAll('editsublet.asp?','sublet','<%=crs("SubletID")%>','','')"><img border=0 src=newimages/edit2.png width=16 height=17></a></font><a onmouseover="javascript:status='Delete Sublet Item';return true" onmouseout="javascript:status='Shop Boss Pro';return true" href="javascript:setCookie('<%=comrs("displayorder")%>');document.body.onunload='';conDel('delroitem.asp?','sublet','<%=crs("SubletID")%>','','')"><img border="0" src="newimages/delete.png" width="16" height="17"></a></td>
  </tr>
  <%
  	ttlsub = ttlsub + crs("subletprice")
  	comttlsub = comttlsub + crs("subletprice")
    crs.movenext
   wend
   %>

   </table>
   <%
  end if
  %>
  <!-- sublet ends -->

      <div style="font-weight:bold;text-align:right;padding:5px;padding-right:20px;">Total This Issue: &nbsp;<%=formatcurrency(comttllbr+comttlprts+comttlsub)%></div>

</td></tr></table>
<%
   comttllbr = 0
   comttlprts = 0
   comttlsub = 0
   comrs.movenext
   cntr = cntr + 1
   if cntr > 1 then
   tcolor = cntr mod 2
    if tcolor = 0 then
  	  mbgc = "#0066CC"
    else
  	  mbgc = "#0066CC"
    end if
   end if
  wend
end if


%>
</div>
