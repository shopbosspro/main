<%
on error resume next
shopname = request("shopname")
'response.write request("invlocation") & "<BR>"
if request("mobile") = "y" then
	invloc = request("invlocation")
	filename = server.mappath(".") & replace(invloc,"/","\")
else
	filename = server.mappath(".") & "\" & replace(request("invlocation"),"/","\")
end if



'response.write filename

set mail = server.createobject("persits.mailsender")
mail.host = "smtp.emailsrvr.com"

mail.port = 587
mail.from = "no-reply@shopbosspro.com"
mail.username = "no-reply@shopbosspro.com"
mail.password = "45ley92!"
mail.addreplyto request("shopemail")
mail.fromname = shopname
mail.addaddress request("customeremail")
mail.addattachment filename
mail.subject = "Copy of your repair order from " & shopname
mb = "This is a copy of the repair order that you electronically signed for " & shopname & ".  Thanks for using our E-Signature System"
mail.body = mb
mail.send

if err.number = 0 then
	response.write "success"
else
	response.write err.number & "-" & err.description
end if
%>
