<!-- #include file=aspscripts/connwoshopid.asp -->
<%

Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function


function validateInfo()

	' verify correct information and return success or failure
	phonevalidate = "no"
	addressvalidate = "no"
	stmt = "select customeraddress ca,customerphone ch,customerwork cw,cellphone cp from repairorders where shopid = '" & shopid & "' and roid = " & roid
	set rs = con.execute(stmt)
	if not rs.eof then
		
		' validate the phone number
		if right(rs("ch"),4) = l4 or right(rs("cw"),4) = l4 or right(rs("cp"),4) = l4 then
			phonevalidate = "yes"
		else
			phonevalidate = "no"
			validateInfo = "phone error"
			exit function
		end if

		'validate the address
		if lcase(rs("ca")) = a then
			addressvalidate = "yes"
		else
			validateInfo = "address error"
			addressvalidate = "no"
			exit function
		end if
		
		if phonevalidate = "yes" and addressvalidate = "yes" then
			validateInfo = "yes|" & rs("ch") & "," & rs("cw") & "," & rs("cp") & "|" & l4 & "|" & lcase(rs("ca")) & "|" & a
		else
			validateInfo = "no"
		end if
	else
		validateInfo = "recordset error"
	end if

end function

on error resume next

shopid = base64decode(request("shopid"))
roid = base64decode(request("roid"))
l4 = request("ph")
a = lcase(request("a"))
sig = request("sig")

x = validateInfo()
'response.write "x:" & x & "<BR>"

if left(x,3) = "yes" then

	'on error resume next

	' valid info, add the customer signature to the RO
	' save the customer validation info for future reference
	' split the function results on the pipe to add to the esig database
	'response.write "validateInfo:" & validateInfo & "<BR>"
	validar = split(validateInfo,"|")
	stmt = "insert into esig (roid,shopid,validatedaddress,actualaddress,validatedphone,actualphone,signaturebox,agreecheck) values (" & roid & ",'" & shopid & "','"
	stmt = stmt & replace(validar(4),"'","''") & "','" & replace(validar(3),"'","''") & "','" & validar(2) & "','" & validar(1) & "','" & replace(sig,"'","''") & "','true')"
	con.execute(stmt)
	
	' now generate the RO again and add the signature
    set fso = server.createobject("scripting.filesystemobject")
    shopfolder = server.mappath(".") & "\pdfinvoices\" & shopid
    if fso.folderexists(shopfolder) then
    	url = "https://" & request.servervariables("SERVER_NAME") & "/sbp/pdfinvoices/" & shopid & "/printpdfro.asp?sig=" & server.urlencode(sig) & "&shopid=" & shopid & "&roid=" & roid
    else
    	url = "https://" & request.servervariables("SERVER_NAME") & "/sbp/pdfinvoices/printpdfro.asp?sig=" & server.urlencode(sig) & "&shopid=" & shopid & "&roid=" & roid
    end if
    'response.write url & "<BR>"
	set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")
	xml.open "GET", url, true
	xml.send()
	xml.waitforresponse
	'response.write xml.readystate
	
	'Turn off error handling
	'On Error Resume Next
	
	'Wait for up to 3 seconds if we've not gotten the data yet
	'do while xml.readyState <> 4
	'	response.write xml.readystate & "<BR>"
	'	xml.waitForResponse 1000
	'loop
	
	strdata = xml.responsetext
	
	newro = server.mappath("." & "\pdfinvoices" & replace(strdata,"/","\"))
	
	' validate the path to the new RO
	'urllocation = replace(request("firstro"),"https://<%=request.servervariables("server_name")%>/sbp","")
	'rolocation = server.mappath("." & replace(urllocation,"/","\"))
	'response.write rolocation
	set fso = server.createobject("scripting.filesystemobject")
	'if fso.fileexists(rolocation) then
	'	fso.deletefile(rolocation)
	'end if
	
	' now copy the file from temp location to permanent upload location
	uploadpath = "d:\savedinvoices\" & shopid
	if fso.folderexists(uploadpath) then
		
		'check for ro folder
		shopuploadpath = uploadpath & "\" & roid
		if fso.folderexists(shopuploadpath) then
			
			' copy the file to that folder
			filenamearray = split(strdata,"/")
			'response.write "fullpath: <br>" & newro & "<BR>" & uploadpath & "\" & roid & "\" & filenamearray(ubound(filenamearray))
			fso.movefile newro, shopuploadpath & "\" & filenamearray(ubound(filenamearray))
			
		else
			fso.createfolder shopuploadpath
			filenamearray = split(strdata,"/")
			'response.write "createrofolder:<br>" & newro & "<BR>" & uploadpath & "\" & roid & "\" & filenamearray(ubound(filenamearray))
			fso.movefile newro, shopuploadpath & "\" & filenamearray(ubound(filenamearray))
		end if
		
	else
		
		fso.createfolder uploadpath
		fso.createfolder uploadpath & "\" & roid
		filenamearray = split(strdata,"/")
		'response.write "createshopandrofolder: <br>" & newro & "<BR>" & uploadpath & "\" & roid & "\" & filenamearray(ubound(filenamearray))
		fso.movefile newro, uploadpath & "\" & roid & "\" & filenamearray(ubound(filenamearray))
	end if
	
	filenamearray = split(strdata,"/")
	response.write "savedinvoices" & "/" & shopid & "/" & roid & "/" & filenamearray(ubound(filenamearray))
	
    if err.number <> 0 then
    	'response.write err.description
    end if
    
else
	response.write "validationerror"
end if

if err.number <> 0 then
	response.write "programerror"
	'response.write err.description
end if
%>
