<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Sub getSolicitacoesColeta()
		Dim sSql, arrSolColeta, intSolColeta, i
		if trim(request.querystring("busca")) <> "" then
				sSql = "select " & _
						"distinct(a.numero_solicitacao_geracao), " & _
						"a.idsolicitacao, " & _
						"b.qtd_cartuchos, " & _
						"a.data_baixa, " & _
						"b.ismaster, " & _
						"b.ismaster, " & _
						"a.idcliente, " & _
						"b.status_coleta_idstatus_coleta,  " & _
						"e.status_coleta " & _
						"from solicitacoes_resgate_ponto as a " & _
						"left join solicitacao_coleta as b " & _
						"on a.idsolicitacao = b.idsolicitacao_coleta " & _
						"left join pontos_coleta as d " & _
						"on a.idcliente = d.idpontos_coleta " & _
						"left join status_coleta as e " & _
						"on b.status_coleta_idstatus_coleta = e.idstatus_coleta where a.numero_solicitacao_geracao = '" & trim(request.querystring("busca")) & "'"
		else
			If CInt(Request.QueryString("StatusSol")) = 0 Then
				sSql = "select " & _
						"distinct(a.numero_solicitacao_geracao), " & _
						"a.idsolicitacao, " & _
						"b.qtd_cartuchos, " & _
						"a.data_baixa, " & _
						"b.ismaster, " & _
						"b.ismaster, " & _
						"a.idcliente, " & _
						"b.status_coleta_idstatus_coleta,  " & _
						"e.status_coleta " & _
						"from solicitacoes_resgate_ponto as a " & _
						"left join solicitacao_coleta as b " & _
						"on a.idsolicitacao = b.idsolicitacao_coleta " & _
						"left join pontos_coleta as d " & _
						"on a.idcliente = d.idpontos_coleta " & _
						"left join status_coleta as e " & _
						"on b.status_coleta_idstatus_coleta = e.idstatus_coleta"
			Else
				If CInt(Request.QueryString("StatusSol")) > 0 Then
				sSql = "select " & _
						"distinct(a.numero_solicitacao_geracao), " & _
						"a.idsolicitacao, " & _
						"b.qtd_cartuchos, " & _
						"a.data_baixa, " & _
						"b.ismaster, " & _
						"b.ismaster, " & _
						"a.idcliente, " & _
						"b.status_coleta_idstatus_coleta,  " & _
						"e.status_coleta " & _
						"from solicitacoes_resgate_ponto as a " & _
						"left join solicitacao_coleta as b " & _
						"on a.idsolicitacao = b.idsolicitacao_coleta " & _
						"left join pontos_coleta as d " & _
						"on a.idcliente = d.idpontos_coleta " & _
						"left join status_coleta as e " & _
						"on b.status_coleta_idstatus_coleta = e.idstatus_coleta where b.status_coleta_idstatus_coleta = " & request.QueryString("StatusSol")
				 end if
			End If
		end if

		Call search(sSql, arrSolColeta, intSolColeta)

		Dim iNumPags
		
		iNumPags = 20 'numero de registros mostrados
		
		Response.Write "<tr><td colspan=2>"
		Response.Write Paginacao(iNumPags, intSolColeta, request("pag"), "frmSolicitacoesResgatePontoColeta", Request.ServerVariables("QUERY_STRING"))
		Response.Write "</td></tr>"

		With Response
			If intSolColeta > -1 Then
				'For i=0 To intSolColeta
				For i = vlrMinimo(iNumPags, intSolColeta, request("pag")) to vlrMaximo(iNumPags, intSolColeta, request("pag"))
					If i Mod 2 = 0 Then
						.Write "<tr>"
						.Write "<td width='5%' align='center' class='classColorRelPar'><img class='imgexpandeinfo' src='img/buscar.gif' alt='Verificar Solicitação de Coleta' onClick=""javascript:window.open('frmeditsolicitacaoresgateponto.asp?idsolic="&arrSolColeta(1,i)&"','','width=500,height=460,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" ></td>"
						.Write 	"<td class='classColorRelPar' align='center'>"&getNomeFantasia(arrSolColeta(6,i))&"</td>"
						.Write 	"<td width='15%' class='classColorRelPar' align='center'>"&arrSolColeta(0,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td align='center' width='15%' class='classColorRelPar'>##/##/####</td>"
						Else
							.Write	"<td align='center' class='classColorRelPar'>"&DateRight(arrSolColeta(3,i))&"</td>"
						End If
						.Write	"<td class='classColorRelPar'  width='15%' align='center'>"&arrSolColeta(2,i)&"</td>"
						.write "<td class='classColorRelPar'  width='15%' align='center'>"&arrSolColeta(8,i)&"</td>"
						.Write "</tr>"
					Else
						.Write "<tr>"
						.Write "<td width='5%' align='center' class='classColorRelImpar'><img class='imgexpandeinfo' src='img/buscar.gif' alt='Verificar Solicitação de Coleta' onClick=""javascript:window.open('frmeditsolicitacaoresgateponto.asp?idsolic="&arrSolColeta(1,i)&"','','width=500,height=460,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" ></td>"
						.Write 	"<td class='classColorRelImpar' align='center'>"&getNomeFantasia(arrSolColeta(6,i))&"</td>"
						.Write 	"<td width='15%' class='classColorRelImpar' align='center'>"&arrSolColeta(0,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td align='center' width='15%' class='classColorRelImpar'>##/##/####</td>"
						Else
							.Write	"<td align='center' class='classColorRelImpar'>"&DateRight(arrSolColeta(3,i))&"</td>"
						End If
						.Write "<td class='classColorRelImpar'  width='15%' align='center'>"&arrSolColeta(2,i)&"</td>"
						.write "<td class='classColorRelImpar'  width='15%' align='center'>"&arrSolColeta(8,i)&"</td>"
						.Write "</tr>"
					End If
				Next
				.Write "<tr><td colspan=2>"
				.Write Paginacao(iNumPags, intSolColeta, request("pag"), "frmSolicitacoesResgatePontoColeta", Request.ServerVariables("QUERY_STRING"))
				.Write "</td></tr>"
			Else
				.Write "<tr>"
				.Write	"<td colspan='6' align='center' class='classColorRelPar'><b>Nenhum Solicitação Encontrada</b></td>"
				.Write "</tr>"
			End If
		End With
	End Sub

	Sub getStatusColeta()
		Dim sSql, arrStatus, intStatus, i
		Dim sSelected
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta] " & _
						"WHERE idStatus_coleta IN (1,2,3,4,6)"
						
		Call search(sSql, arrStatus, intStatus)
		With Response
			If intStatus > -1 Then
				For i=0 To intStatus
					If Request.QueryString("StatusSol") = CStr(arrStatus(0,i)) Then
						sSelected = "selected"
					Else
						sSelected = ""
					End If
					.Write "<option value='"&arrStatus(0,i)&"' "&sSelected&">"&arrStatus(1,i)&"</option>"
				Next
			End If
		End With
	End Sub

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		Dia = Left(sData, 2)
		Dia = Replace(Dia, "/", "")
		If Len(Dia) = 1 Then
			Dia = "0" & Dia
		End If
		If Len(Replace(Left(sData, 2), "/", "")) = 1 Then
			Mes = Mid(sData, 3, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		Else
			Mes = Mid(sData, 4, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		End If
		Ano = Right(sData, 4)
		Ano = Replace(Ano, "/", "")
		If Len(Ano) = 1 Then
			Ano = "0" & Ano
		End If
		DateRight = Mes & "/" & Dia & "/" & Ano
	End Function

	function getNomeFantasia(id)
		dim sql, arr, intarr, i
		sql = "select nome_fantasia from pontos_coleta where idpontos_coleta = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getNomeFantasia = arr(0,i)
			next
		else
			getNomeFantasia = ""
		end if
	end function

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="" name="frmOperacionalAdm" method="POST">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td colspan="4" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="4" id="explaintitle" align="left">
								Número Solicitação: <input type="text" name="busca" class="text" value="<%= request.querystring("busca") %>" size="40" />
								<input type="button" name="btnprocurar" value="Procurar" class="btnform" onClick="window.location.href='frmsolicitacoesresgatepontocoleta.asp?busca=' + document.frmOperacionalAdm.busca.value + ''" />
							</td>
						<tr>
							<td colspan="2" id="explaintitle" align="center">Solicitações de Resgate</td>
							<td id="explaintitle" align="right">
								Status :
								<select name="cbStatusColeta" class="select" onChange="window.location.href='frmsolicitacoesresgatepontocoleta.asp?StatusSol=' + this.value;">
									<option value="0">Todas</option>
									<%Call getStatusColeta()%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="5">
								<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
									<tr>
										<th>Ações</th>
										<th>Nome Fantasia</th>
										<th>N° Solicitação</th>
										<th>DT. Baixa</th>
										<th>Quantidade</th>
										<th>Status</th>
									</tr>
									<%Call getSolicitacoesColeta()%>
								</table>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</form>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
