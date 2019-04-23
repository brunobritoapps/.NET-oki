<!--#include file="../_config/_config.asp" -->
<%
Call open()

Dim lId, sTipo
Dim sqlGrupo

lId = request("Id")
sTipo = request("Tipo")

if sTipo = "A" then
	sqlGrupo = "UPDATE Solicitacao_coleta SET Status_coleta_idStatus_coleta = 2 WHERE idSolicitacao_coleta = " & lId
else
	sqlGrupo = "UPDATE Solicitacao_coleta SET Status_coleta_idStatus_coleta = 3 WHERE idSolicitacao_coleta = " & lId
end if

oConn.execute(sqlGrupo)

Call open()

Response.Write "<script>"
Response.Write "window.opener.location.reload();"
Response.Write "window.close()"
Response.Write "</script>"
%>