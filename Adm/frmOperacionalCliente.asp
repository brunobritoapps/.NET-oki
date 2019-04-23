<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%
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

	Sub getSolicitacoesColeta()
		Dim sSql, arrSolColeta, intSolColeta, i

		If Request.QueryString("StatusSol") = "" Or Request.QueryString("StatusSol") = 0 Then
			sSql = "SELECT " & _
							"[A].[idSolicitacao_coleta], " & _
							"[A].[numero_solicitacao_coleta], " & _
							"[A].[qtd_cartuchos], " & _
							"[A].[data_recebimento], " & _
							"[A].[motivo_status], " & _
							"[A].[isMaster] " & _
							"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _
							"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _
							"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _
							"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
							"ON [B].[Clientes_idClientes] = [C].[idClientes] " & _
							"WHERE " & _
							"[B].[Clientes_idClientes] = "&Session("IDCliente")&" OR C.[cod_cli_consolidador] = "&Session("IDCliente")
		Else
			sSql = "SELECT " & _
							"[A].[idSolicitacao_coleta], " & _
							"[A].[numero_solicitacao_coleta], " & _
							"[A].[qtd_cartuchos], " & _
							"[A].[data_recebimento], " & _
							"[A].[motivo_status], " & _
							"[A].[isMaster] " & _
							"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _
							"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _
							"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _
							"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
							"ON [B].[Clientes_idClientes] = [C].[idClientes] " & _
							"WHERE " & _
							"[A].[Status_coleta_idStatus_coleta] = " & Request.QueryString("StatusSol") & " " & _
							"AND [B].[Clientes_idClientes] = "&Session("IDCliente")&" OR C.[cod_cli_consolidador] = "&Session("IDCliente")

		End If

		Call search(sSql, arrSolColeta, intSolColeta)

		With Response
			If intSolColeta > -1 Then
				For i=0 To intSolColeta
					If i Mod 2 = 0 Then
						.Write "<tr style=""cursor:pointer;"" onClick=""javascript:window.open('frmViewSol.asp?idsol="&arrSolColeta(0,i)&"','','width=720,height=600,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"
						.Write 	"<td class='classColorRelPar'>"&arrSolColeta(1,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td class='classColorRelPar'>Não Disponível</td>"
						Else
							if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
								.Write	"<td class='classColorRelPar'>"&DateRight(arrSolColeta(3,i))&"</td>"
							else
								.Write	"<td class='classColorRelPar'>"&arrSolColeta(3,i)&"</td>"
							end if
						End If
						.Write	"<td class='classColorRelPar'>"&arrSolColeta(2,i)&"</td>"
						.Write "</tr>"
					Else
						.Write "<tr style=""cursor:pointer;"" onClick=""javascript:window.open('frmViewSol.asp?idsol="&arrSolColeta(0,i)&"','','width=720,height=600,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"
						.Write 	"<td class='classColorRelImpar'>"&arrSolColeta(1,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td class='classColorRelImpar'>Não Disponível</td>"
						Else
							if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
								.Write	"<td class='classColorRelImpar'>"&DateRight(arrSolColeta(3,i))&"</td>"
							else
								.Write	"<td class='classColorRelImpar'>"&arrSolColeta(3,i)&"</td>"
							end if
						End If
						.Write	"<td class='classColorRelImpar'>"&arrSolColeta(2,i)&"</td>"
						.Write "</tr>"
					End If
				Next
			Else
				.Write "<tr>"
				.Write	"<td colspan='3' align='center'>Nenhum Solicitação Encontrada</td>"
				.Write "</tr>"
			End If
		End With
	End Sub

	function getSolicitacoesResgate()
		dim sql, arr, intarr, i
		dim arr2, intarr2, j
		dim html, style
		sql = "SELECT distinct([numero_solicitacao_geracao]) " & _
			  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
			  "WHERE [idcliente] = " & session("IDCliente")
		call search(sql, arr2, intarr2)
		if intarr2 > -1 then
			for j=0 to intarr2
				sql = "SELECT [numero_solicitacao_geracao] " & _
						  ",[idSolicitacoes_resgate] " & _
						  ",[data_solicitacao_resgate] " & _
						  ",[idcliente] " & _
						  ",[quantidade] " & _
						  ",[idsolicitacao] " & _
					  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
					  "WHERE [numero_solicitacao_geracao] = '" & arr2(0,j)&"'"
'				response.write sql
				call search(sql, arr, intarr)
				if intarr > -1 then
					for i=0 to intarr
						if i mod 2 = 0 then
							style="class=""classColorRelPar"" style=""cursor:pointer;"""
						else
							style="class=""classColorRelImpar"" style=""cursor:pointer;"""
						end if
						html = html & "<tr>"
						html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(5,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(0,i)&"</td>"
						if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
							html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(5,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&DateRight(arr(2,i))&"</td>"
						else
							html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(5,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(2,i)&"</td>"
						end if
						html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(5,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(4,i)&"</td>"
						html = html & "</tr>"
					next
				else
					html = html & "<tr>"
					html = html & "<td colspan=""3"" "&style&"><b>Nenhum registro encontrado</b></td>"
					html = html & "</tr>"
				end if
			next
		end if
		getSolicitacoesResgate = html
	end function

	Sub getStatusColeta()
		Dim sSql, arrStatus, intStatus, i
		Dim sSelected
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta]"
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

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td colspan="3" id="explaintitle" align="center">Painel de Controle</td>
						</tr>
						  <% If not isnull(session("IDCliente")) and not isempty(session("IDCliente")) Then %>
								<tr>
									<td id="explaintitle" align="right" colspan="3" style="padding:4px 4px 4px 4px;">
										<a href="frmoperacionalcliente.asp?logoff=true" style="color: #FFFFFF;">Logoff do Sistema</a>
									</td>
								</tr>
						  <% End If %>
						<tr>
							<%If Session("isMaster") = 1 Then%>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" src="img/cpanel.png" width="32" height="32" alt="Operacional Cadastro [Atualize as informações sobre a Empresa]" onClick="window.location.href='frmEditaCadastroCliente.asp'" /><br />
								<a href="frmEditaCadastroCliente.asp" class="linkOperacional">Alterar Cadastro</a>
							</td>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" src="img/contato.png" alt="Adicionar Contato [Adicione um novo contato para efetuar Solicitações de Coleta]" onClick="window.location.href='frmAddContato.asp'" /><br />
								<a href="frmAddContato.asp" class="linkOperacional">Manutenção de Contatos</a>
							</td>
							<%End If%>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Nova Solicitação de Coleta [Envie uma Nova Solicitação de Coleta]" onClick="window.location.href='frmAddSolicitacao.asp'" /><br />
								<a href="frmAddSolicitacao.asp" class="linkOperacional">Nova Solicitação de Coleta</a>
							</td>
						</tr>
						<tr>
						<%if Session("isColetaDomiciliar") = 1 and Session("isMaster") = 1 and session("cod_cli_consolidador") = 0 and session("cod_bonus") <> "" then%>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" width="38" height="38" src="adm/img/bonus.gif" alt="Bônus Gerado" onClick="javascript:window.open('frmviewbonuscliente.asp','','width=750,height=650,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /><br />
								<a href="frmviewbonuscliente.asp" class="linkOperacional">Bônus Gerado</a>
							</td>
						<%end if%>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" width="38" height="38" src="adm/img/kardex.jpg" alt="Relatórios" onClick="javascript:window.location.href='frmtiporelatorio.asp';" /><br />
								<a href="frmtiporelatorio.asp" class="linkOperacional">Relatórios</a>
							</td>
						<%if Session("isColetaDomiciliar") = 1 and Session("isMaster") = 1 and session("cod_cli_consolidador") = 0 and session("cod_bonus") <> "" then%>
							<td align="center" width="33%">
								<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Nova Solicitação de Coleta [Envie uma Nova Solicitação de Coleta]" onClick="window.location.href='frmAddSolicitacao.asp'" /><br />
								<a href="frmAddSolicitacao.asp" class="linkOperacional">Nova Solicitação de Coleta</a>
							</td>
						<%end if%>
						</tr>
						<tr>
							<td colspan="2" id="explaintitle" align="center">Solicitações de Coleta</td>
							<td id="explaintitle" align="right">
								<select name="cbStatusColeta" class="select" onChange="window.location.href='frmOperacionalCliente.asp?StatusSol=' + this.value;">
									<option value="0">Todas</option>
									<%Call getStatusColeta()%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="3">
								<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
									<tr>
										<th>N° Solicitação</th>
										<th>DT. Recebimento</th>
										<th>Quantidade</th>
									</tr>
									<%Call getSolicitacoesColeta()%>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="3" id="explaintitle" align="center">Solicitações de Resgate</td>
						</tr>
						<tr>
							<td colspan="3">
								<%if Session("isColetaDomiciliar") = 1 and Session("isMaster") = 1 and session("cod_cli_consolidador") = 0 and session("cod_bonus") <> "" then%>
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
										<tr>
											<th>N° Solicitação</th>
											<th>DT. Solic. Resgate</th>
											<th>Quantidade</th>
										</tr>
										<%=getSolicitacoesResgate()%>
									</table>
								<%end if%>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
