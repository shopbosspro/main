<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/conn.asp -->
<%
if len(request("ponumber")) > 0 then
	stmt = "insert into po (shopid,ponumber,`desc`,issuedate,`status`,saletype,salenumber,issuedto,ordertype) values ('" & request.cookies("shopid") & "', " & request("ponumber") _
	& ", '" & replace(request("desc"),"'","''") & "', '" & dconv(request("issuedate")) & "', '" & request("status") & "', 'RO','" & request("roid") & "', '" & replace(request("issuedto"),"'","''") _
	& "', 'RO')"
	'response.write stmt
	con.execute(stmt)
	stmt = "select id,issuedto from po where shopid = '" & request.cookies("shopid") & "' and ponumber = " & request("ponumber")
	set rs = con.execute(stmt)
	poid = rs("id")
%>
<script type="text/javascript">
	
	if(parent.document.getElementById("ponumber")){
		var select = parent.document.getElementById("ponumber"); 
		select.options[select.options.length] = new Option('<%=request("ponumber") & " - " & rs("issuedto")%>', '<%=poid%>')
		var searchfor = "<%=request("ponumber") & " - " & rs("issuedto")%>"
	
	    for(var i = 0, j = select.options.length; i < j; ++i) {
	        if(select.options[i].innerHTML === searchfor) {
	           select.selectedIndex = i;
	           break;
	        }
	    }
		
	}else{
		var dropbox
		dropbox = "<select id='ponumber' name='ponumber'><option value='<%=poid%>'><%=request("ponumber") & " - " & rs("issuedto")%></option></select>"	
		parent.document.getElementById("pospan").innerHTML = dropbox
	}

	parent.document.theform.ponumber.focus()
	parent.document.getElementById("newpo").style.display = 'none'
	parent.document.getElementById("hider").style.display = 'none'

</script>
<%
end if
%>
<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Add New PO To RO #</title>
<style type="text/css">
body{
	font-family:Arial, Helvetica, sans-serif;
	margin:0px;
}
.style1 {
	text-align: right;
}
.style2 {
	text-align: center;
}
.style3 {
	color: #FFFFFF;
}
.style4 {
	font-size: large;
}
.style5 {
	color: #FF0000;
}
</style>

<script type="text/javascript">

function selectAll(){

  var checkboxes = new Array(); 
  checkboxes = document.getElementsByTagName('input');
 
  for (var i=0; i<checkboxes.length; i++)  {
    if (checkboxes[i].type == 'checkbox')   {
      checkboxes[i].checked = true;
    }
  }
}

function deselectAll(){

  var checkboxes = new Array(); 
  checkboxes = document.getElementsByTagName('input');
 
  for (var i=0; i<checkboxes.length; i++)  {
    if (checkboxes[i].type == 'checkbox')   {
      checkboxes[i].checked = false;
    }
  }
}
</script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
<script>

  $(function() {
    var availableTags = [
      <%
      stmt = "select ucase(suppliername) as sn from supplier where shopid = '" & request.cookies("shopid") & "' order by suppliername"
      set rs = con.execute(stmt)
      if not rs.eof then
      	do until rs.eof
      		tags = tags &  chr(34) & rs("sn") & chr(34) & ","
      		rs.movenext
      	loop
      else
      	response.write "NONE"
      end if
      tags = left(tags,len(tags)-1)
      response.write tags
      %>

    ];
    $( "#issuedto" ).autocomplete({
      source: availableTags
    });
  });

  </script>



<link rel="STYLESHEET" type="text/css" href="css/rich_calendar.css"/>
<script language="JavaScript" type="text/javascript" src="javascripts/rich_calendar.js"></script>
<script language="JavaScript" type="text/javascript" src="javascripts/rc_lang_en.js"></script>
<script language="javascript" type="text/javascript" src="javascripts/domready.js"></script>
<script type="text/javascript" src="javascripts/killenter.js" ></script>
<script type="text/javascript">
var cal_obj2 = null;

var format = '%m/%d/%Y';

// show calendar
function show_cal(el,a) {

	if (cal_obj2) return;
	document.getElementById("dfield").value = a
	var text_field = document.getElementById(document.getElementById("dfield").value);
	text_field.disabled = true;
	cal_obj2 = new RichCalendar();
	cal_obj2.start_week_day = 0;
	cal_obj2.show_time = false;
	cal_obj2.language = 'en';
	cal_obj2.user_onchange_handler = cal2_on_change;
	cal_obj2.user_onclose_handler = cal2_on_close;
	cal_obj2.user_onautoclose_handler = cal2_on_autoclose;

	cal_obj2.parse_date(text_field.value, format);

	cal_obj2.show_at_element(text_field, "adj_right-");
	//cal_obj2.change_skin('alt');

}

// user defined onchange handler
function cal2_on_change(cal, object_code) {
	if (object_code == 'day') {
		document.getElementById(document.getElementById("dfield").value).disabled = false;
		document.getElementById(document.getElementById("dfield").value).value = cal.get_formatted_date(format);
		
		cal.hide();
		cal_obj2 = null;
	}
}

// user defined onclose handler (used in pop-up mode - when auto_close is true)
function cal2_on_close(cal) {
		document.getElementById(document.getElementById("dfield").value).disabled = false;
		cal.hide();
		cal_obj2 = null;

}

// user defined onautoclose handler
function cal2_on_autoclose(cal) {
	document.getElementById(document.getElementById("dfield").value).disabled = false;
	cal_obj2 = null;
}

function cancelNewPO(){

	parent.document.getElementById("newpo").style.display = 'none'
	parent.document.getElementById("hider").style.display = 'none'

}

</script>

</head>

<%
stmt = "select ponumber from po where shopid = '" & request.cookies("shopid") & "' order by ponumber desc limit 1"
set rs = con.execute(stmt)
if not rs.eof then
	newponumber = cdbl(rs("ponumber")) + 1
else
	newponumber = 1000
end if
%>

<body>
<form method="post" action="addpartfoundnewpo.asp">
	<input name="roid" type="hidden" value="<%=request("roid")%>" />
<input type="hidden" name="dfield" id="dfield"/>
<div style="height: 46px; background-image:url('newimages/wipheader.jpg');padding-top:20px;" class="style2">
	<strong><span class="style2"><span class="style3"><span class="style4">Add New PO To RO #<%=request("roid")%></span></span></span></strong></div>
<table align="center" cellpadding="3" cellspacing="0" style="width: 80%">
	<tr>
		<td class="style1"><strong>PO Number</strong></td>
		<td class="style5"><strong><%=newponumber%><input name="ponumber" type="hidden" value="<%=newponumber%>" />
		</strong>&nbsp;</td>
		<td class="style1"><strong>Issue Date</strong></td>
		<td><input onfocus="show_cal(this,'issuedate');" id="issuedate" type="text" value="<%=date%>" name="issuedate" />&nbsp;</td>
	</tr>
	<tr>
		<td class="style1"><strong>Status</strong></td>
		<td>
		<select name="status">
			
			<option <%=py%> value="Pending">Pending</option>
			
			<option <%=iy%> value="Issued">Issued</option>
			<option <%=cy%> value="Closed">Closed</option>
		</select>
		&nbsp;</td>
		<td class="style1"><strong>Issued To</strong>&nbsp; </td>
		<td><div class="ui-widget">

 <input style="font-size:small; width: 203px;" id="issuedto" name="issuedto" />

</div>

</td>
	</tr>
	<tr>
		<td class="style1"><strong>Comments</strong></td>
		<td colspan="3">
		<textarea name="desc" rows="2" style="width: 570px"></textarea></td>
	</tr>
	<tr>
		<td class="style2" colspan="4"><br />
		<input name="Button1" type="submit" value="Add PO" />
		<input onclick="cancelNewPO()" name="Button2" type="button" value="Cancel" /></td>
	</tr>
</table>
<hr/>
</form>
</body>

</html>
