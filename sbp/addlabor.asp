<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=aspscripts/adovbs.asp -->
<%
if len(request("labor")) > 0 then
	'verify all information
	'connect to database and update information
	srostmt = "select LaborID from labor where shopid = '" & request.cookies("shopid") & "' order by LaborID DESC"
	set rs = con.execute(srostmt)
	if not rs.eof then
		rs.movefirst
		newid = cdbl(rs("LaborID")) + 1
	else
		newid = 1
	end if
	set rs = nothing
	mytech = request("tech")
	'response.write "my tech = " & mytech
	if len(request("comid")) > 0 then
		comid = request("comid")
	else
		comid = 0
	end if

	'get employee hourly rate
	'create array of first and last name
	marray = split(request("tech"), ",")
	rstmt = "select hourlyrate from employees where shopid = '" & request.cookies("shopid") & "' and employeelast = '" & trim(marray(0)) & "' and employeefirst = '" & trim(marray(1)) & "'"
	set rrs = con.execute(rstmt)
	if not rrs.eof then
		rrs.movefirst
		ehr = rrs("hourlyrate")
	end if
	stmt = "insert into labor (shopid, roid, hourlyrate, laborhours, labor, tech, linetotal,discount,discountpercent, complaintid, techrate) values ('" & request.cookies("shopid") & "', "
	'response.write stmt
	stmt=stmt& request("roid")
	if isnumeric(request("rate")) then
		hrate = request("rate")
	else
		hrate = 0
	end if
	if isnumeric(request("hours")) then
		hhours = request("hours")
	else
		hhours = 0
	end if
	stmt=stmt&", " & hrate
	stmt=stmt&", " & hhours
	stmt=stmt&", '" & replace(ucase(request("Labor")),"'","''")
	stmt=stmt&"', '" & ucase(mytech)
	' *******  calculate the discount
	linettl = request("calcrate")
	stmt=stmt&"', " & linettl
	stmt=stmt&"," & request("discountdollars")
	stmt=stmt&"," & request("discount")
	stmt=stmt&", " & comid
	stmt=stmt&", " & ehr * request("hours") & ")"
	mysplit = "no"

	'response.write "<BR>" & stmt
	con.execute stmt
	
	recordAudit "Add Labor", "Added Labor " & ucase(replace(ucase(request("Labor")),"'","''")) & " (" & hhours & " hours) to RO#" & request("roid")
	
	if request("addanother") = "Yes" then
		'response.redirect "addlabor.asp?comid=" & request("comid") & "&roid=" & request("roid")
	else
		response.redirect "ro.asp?roid=" & request("roid")
	end if
end if

strsql = "Select EmployeeLast, EmployeeFirst from employees where shopid = '" & request.cookies("shopid") & "' and Active = 'yes' and showtechlist = 'yes' order by employeelast asc"
set rs = con.execute(strsql)
%>
<!DOCTYPE html>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252"/>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0"/>
<meta name="ProgId" content="FrontPage.Editor.Document"/>
<link type="text/css" href="jquery/jquery-ui.css" rel="stylesheet" />
<script type="text/javascript" src="jquery/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="jquery/jquery-ui.js"></script>


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
.style2 {
	font-weight: bold;
}
.style6 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: x-small;
}
.style9 {
	border: 1px solid #000000;
}
.style11 {
	font-weight: bold;
	text-align: center;
	color: #FFFFFF;
		background-image: url('newimages/wipheader.jpg');
}
.style12 {
	text-align: center;
	font-size: small;
}
.style13 {
	color: #B90505;
}
.style14 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
}
.style15 {
	font-size: small;
}

#addldesc, #canceldesc{
	padding:15px;
	cursor:pointer;
	background-color:navy;
	color:white;
	border-radius:5px;
}
#addldesc:hover, #canceldesc:hover{
	padding:15px;
	cursor:pointer;
	background-color:maroon;
	color:white;
	border-radius:5px;
}


-->
</style>
		<script type="text/javascript">
		
         
		function fnIsAppleMobile() 
		{
		    if (navigator && navigator.userAgent && navigator.userAgent != null) 
		    {
		        var strUserAgent = navigator.userAgent.toLowerCase();
		        var arrMatches = strUserAgent.match(/(iphone|ipod|ipad)/);
		        if (arrMatches != null) {
		             return true;
		        }
		    } // End if (navigator && navigator.userAgent) 
		
		    return false;
		}
		var mapp = fnIsAppleMobile()
		//alert(mapp)
			$(function(){

				//hover states on the static widgets
				$('#dialog_link, ul#icons li').hover(
					function() { $(this).addClass('ui-state-hover'); },
					function() { $(this).removeClass('ui-state-hover'); }
				);

			});

			$(function() {
				var availableTags = [
					<%
					stmt = "select * from repairs where shopid = '" & request.cookies("shopid") & "' order by repair"
					'response.write stmt & chr(10)
					set repairrs = con.execute(stmt)
					if not repairrs.eof then
						do until repairrs.eof
							response.write "'" & repairrs("repair") & "'," & chr(10)
							repairrs.movenext
						loop
					end if
					
					stmt = "select * from repairs where shopid = '0' order by repair"
					set repairrs = con.execute(stmt)
					do until repairrs.eof
						response.write "'" & repairrs("repair") & "'," & chr(10)
						repairrs.movenext
					loop
					
					%>
				];
				if(!mapp){
					$( "#autocomplete" ).autocomplete({
						source: availableTags
					});
				}
			});

		</script>

<script language="JavaScript">

	function opensplit(){
		//check for values in labor and hours
		if(document.theform.Labor.value.length==0){
			alert("You must enter the labor item first")
			okframe = "no"
			return
		}else{
			okframe = "yes"
		}
		if(document.theform.hours.value.length==0){
			alert("You must enter the total number of hours first")
			okframe = "no"
			return
		}else{
			okframe = "yes"
		}
		if(okframe=="yes"){
		document.getElementById("ifname").src = "splittech.asp"
		document.getElementById("ifname").style.visibility="visible"
		
		}
	}
function checkForm(){

	if(!isNumber(document.theform.discount.value.length)){
		document.theform.discount.value = 0
	}
	
	if(document.theform.Labor.value.length == 0){
		alert("Labor Item is required")
		return
	}
	if(!isNumber(document.theform.hours.value)){
		alert("# Hours is required")
		return
	}

	if(!isNumber(document.theform.rate.value)){
		alert("Hourly Rate is required")
		return
	}

	if(!isNumber(document.theform.calcrate.value)){
		alert("Calculated Labor is required")
		return
	}


	document.theform.submit();

}

function isNumber(num) {   
	return (typeof num == 'string' || typeof num == 'number') && !isNaN(num - 0) && num !== ''; 
}

function pd(){
	if (document.theform.discount.value >= 0){
		mypd = ((document.theform.hours.value*document.theform.rate.value) * document.theform.discount.value) / 100
		mypd = Math.round(mypd*100)/100
		mynetpd = Math.round(((document.theform.hours.value*document.theform.rate.value) - mypd)*100)/100
		document.theform.calcrate.value = mynetpd
		document.getElementById("discountconverted").innerHTML = "Discount: $" + mypd
		document.theform.discountdollars.value = mypd
		//document.theform.TotalPrice.value = document.theform.Net.value*document.theform.Quantity.value
	}else{

		if(document.theform.discount.value.length == 0){
			document.theform.discount.value = 0
			document.theform.calcrate.value=document.theform.hours.value*document.theform.rate.value
		}
	}
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
        <div class="style12">
        	<strong>Begin typing your labor description.&nbsp; 
			You will be given a drop down list to choose from or you can 
			continue to type and ignore the list
			<% if request("err") = "yes" then %>
			<br><br><span class="style13">
			Please verify your information below.&nbsp; We were not able to add 
			this to your Repair Order</span>
			<% end if %>
			</strong>
			
			</div>
        <div align="center">
         <center>
         <form method="POST" action="addlabornow.asp" name="theform" id="theform">
         <input type=hidden name=stech value="">
         <input type=hidden name=percent value="">
         <input type=hidden name=split value="no">
          <table width="70%" cellspacing="0" cellpadding="2" class="style9">
           <tr>
            <td class="style11" colspan="2" style="width: 100%; height: 50px">
			Add a Labor Item</td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2">Labor Item</td>
            <td width="70%">
			<input id="autocomplete" type="text" name="Labor" size="63" style="font-size:18px;font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:350px" value="<%=request("Labor")%>" class="style14">&nbsp; 
			<input id="opener" name="button" type="button" value="Add New Labor Description to the List"></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2"># Hours</td>
            <td width="70%">
			<input  onblur="javascript:this.style.backgroundColor='white';document.theform.calcrate.value=document.theform.hours.value*document.theform.rate.value" onfocus="javascript:this.style.backgroundColor='yellow'" type="text" name="hours" size="11" style="font-size:18px;font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" value="<%=request("hours")%>" class="style14"></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2">Technician</td>
            <iframe frameborder=0 width=660 style="visibility:hidden; position:absolute; top:170; left:232" height=500 src='', id=ifname name=ifname></iframe></div>
            <td width="70%" class="style6">
			<div id=tech name=tech class="style6"><select  onblur="javascript:this.style.backgroundColor='white'" onfocus="javascript:this.style.backgroundColor='yellow'" size="1" name="tech">
              
                      <%
                      if not rs.eof then
                      	rs.movefirst
                      	while not rs.eof
                      		response.write "<option value='" & rs("EmployeeLast") & ", " & rs("EmployeeFirst")
                      		response.write "'>" & rs("EmployeeLast") & ", " & rs("EmployeeFirst") & "</option>"
                      		rs.movenext
                      	wend
                      else
                      	needtech = "yes"
                      end if
                      set rs = nothing
                      set rs = con.execute("select HourlyRate,hourlyrate2,hourlyrate3,hourlyrate4,hourlyrate5,hourlyrate6,hourlyrate1label,hourlyrate2label,hourlyrate3label,hourlyrate4label,hourlyrate5label,hourlyrate6label from company where shopid = '" & request.cookies("shopid") & "'")
                      if not rs.eof then
                      	rs.movefirst
                      	hrate = rs("HourlyRate")
                      	hrate2 = rs("HourlyRate2")
                      	hrate3 = rs("HourlyRate3")
                      	hrate4 = rs("HourlyRate4")
                      	hrate5 = rs("HourlyRate5")
                      	hrate6 = rs("HourlyRate6")
                      	hlabel1 = rs("hourlyrate1label")
                      	hlabel2 = rs("hourlyrate2label")
                      	hlabel3 = rs("hourlyrate3label")
                      	hlabel4 = rs("hourlyrate4label")
                      	hlabel5 = rs("hourlyrate5label")
                      	hlabel6 = rs("hourlyrate6label")

                      end if
                      set rs = nothing
                      %>
             </select><%if needtech = "yes" then response.write "<span style='color:red;font-weight:bold;font-size:large'>YOU MUST HAVE AT LEAST ONE TECHNICIAN TO ADD LABOR</span>"%></div></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2">Rate</td>
            <td width="70%">
            <%
            if len(hlabel2) > 0 then
            %>
            <select style="text-transform:uppercase" onchange="document.theform.calcrate.value=document.theform.hours.value*document.theform.rate.value" name="rate" class="style15">
				<option value="<%=hrate%>"><%=hlabel1 & " - " & hrate%></option>
				<%
				if len(hlabel2) > 0 then
				%>
				<option value="<%=hrate2%>"><%=hlabel2 & " - " & hrate2%></option>
				<%
				end if
				if len(hlabel3) > 0 then
				%>
				<option value="<%=hrate3%>"><%=hlabel3 & " - " & hrate3%></option>
				<%
				end if
				if len(hlabel4) > 0 then
				%>
				<option value="<%=hrate4%>"><%=hlabel4 & " - " & hrate4%></option>
				<%
				end if
				if len(hlabel5) > 0 then
				%>
				<option value="<%=hrate5%>"><%=hlabel5 & " - " & hrate5%></option>
				<%
				end if
				if len(hlabel6) > 0 then
				%>
				<option value="<%=hrate6%>"><%=hlabel6 & " - " & hrate6%></option>
				<%
				end if
				%>

				
			</select>
			<%
			else
			%>
			<input value="<%=hrate%>" onblur="javascript:this.style.backgroundColor='white';document.theform.calcrate.value=document.theform.hours.value*document.theform.rate.value" onfocus="javascript:this.style.backgroundColor='yellow'" type="text" name="rate" size="11" style="font-size:18px;font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" class="style14"/>
           <%
           end if
           %>
           </td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2">Discount %</td>
            <td width="70%">
            <input onkeyup="pd()" onblur="javascript:this.style.backgroundColor='white';if(this.value.length==0){this.value=0}" onfocus="javascript:this.style.backgroundColor='yellow'" type="text" name="discount" size="11" style="font-size:18px;font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:122px" class="style14" value="0">&nbsp;&nbsp;&nbsp;
           <span style="color:red;font-weight:bold" id="discountconverted"></span><input type="hidden" id="discountdollars" name="discountdollars" value="0"/></td>
           </tr>
           <tr>
            <td width="30%" align="right" class="style2">Calculated Labor</td>
            <td width="70%">
			<input type="text" name="calcrate" size="11" style="font-size:18px;font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" class="style14"/></td>
           </tr>
           <tr>
            <td width="100%" colspan="2">
             <p align="center">
			 <input type="submit" id="addlaborbutton" onclick="checkForm()" value="Add Labor" class="obutton" style="cursor:pointer"/>
			 <input type="button" onclick="document.theform.addanother.value='Yes';checkForm()" value="Save &amp; Add Another" class="obutton" style="cursor:pointer"/>
			 <input type="button" onclick="location.href='ro.asp?roid=<%=request("roid")%>'" value="Cancel" class="obutton" style="cursor:pointer"> 
			 </td>
           </tr>
          </table>
          <input type="hidden" name="roid" value="<%=request("roid")%>"/><input type="hidden" name="addanother" value="No"/>
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
<div id="dialog-1" title="Add New Labor Description">
	<div style="padding:20px;">Enter a labor description to add to the list</div>
	<div><input name="Text1" id="ldesc" type="text" /></div><br><br>
	<span id="addldesc" >Add Labor Description</span>
	<span id="canceldesc" >Cancel</span>
</div>

<script language="javascript">

 $(function() {
    $( "#dialog-1" ).dialog({
       autoOpen: false,  
       width: 450,
       height: 350,
       modal: true
    });
    
    $( "#opener" ).click(function() {
       $( "#dialog-1" ).dialog( "open" );
    });
    
    $('#canceldesc').click(function(){
    	$('#dialog-1').dialog("close")
    });
    
    $('#addldesc').click(function(){
    
    	$.ajax({
    		data: "shopid=<%=request.cookies("shopid")%>&ld=" + $('#ldesc').val(),
    		url: "php/php/addldesc.php",
    		success: function(r){
    			document.theform.Labor.value = $('#ldesc').val()
    			$('#dialog-1').dialog("close")
    		}
    	});
    	
    });
    
 });


document.theform.Labor.focus()

$('theform').keypress( function( e ) {
  var code = e.keyCode || e.which;

  if( code === 13 ) {
    e.preventDefault();
    return false; 
  }
})
</script>
<p>
&nbsp;</p>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>