<!--#include file="../_config/_config.asp" -->
<%
Call open()

Dim lId, lGrupo
Dim sql, i, v

lId = request("chkIdClientes")
lGrupo = request("cmbAgrupar")

v = split(lId, ",")

for i=0 to ubound(v)
	sql = "UPDATE Clientes SET Grupos_idGrupos = " & lGrupo & " WHERE idClientes = " & v(i)
	oConn.execute(sql)
next

Call open()

Response.Write "<script>"
Response.Write "window.location='frmAgrupamentosAdm.asp?IdGrupos=" & lGrupo & "';"
Response.Write "</script>"
%>