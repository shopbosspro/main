<!-- #include file=aspscripts/conn.asp -->

<%
cr = chr(10)
make = replace(request("make")," ","")
modeldv = modeldv & "var " & make & " = new Array("
stmt = "select distinct model, submodel, make from models where make = '" & request("make") & "' order by model, submodel"
set rs = conn.execute(stmt)
rs.movefirst
cntr = 0
do until rs.eof
	modeldv = modeldv & "'" & trim(rs("model")) & " " & trim(rs("submodel")) & "',"
	rs.movenext
loop
modeldv = left(modeldv,len(modeldv)-1)
modeldv = modeldv & ");" & cr
response.write modeldv
set fso = server.createobject("scripting.filesystemobject")
pname = server.mappath("fpdb/")
pname = pname & "\models.js"
set mfile = fso.opentextfile(pname,8)
mfile.write modeldv
mfile.close
set fso = nothing
%>
