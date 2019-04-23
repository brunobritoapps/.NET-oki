<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
	dim numcoleta
	numcoleta = request.form("txtNumColeta")
	
	sub getAcompanhaSolicitacao()
	dim sql, arr, intarr, i

	if len(trim(numcoleta)) > 0 then
		sql = "SELECT " & _
					 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
					 "B.[numero_solicitacao_coleta], " & _
					 "B.[data_solicitacao] " & _
					 "FROM [marketingoki2].[dbo].[solicitacao_coleta_has_pontos_coleta] AS A " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
					 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
					 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
					 " and B.[numero_solicitacao_coleta] = '" & numcoleta & "' " & _
					 " and B.[isMaster] = 1 and left(B.[numero_solicitacao_coleta],1) <> 'E' " & _
					 " order by B.[numero_solicitacao_coleta] "
'					 " and B.[Status_coleta_idStatus_coleta] <> 3 and B.[Status_coleta_idStatus_coleta] <> 4 and B.[Status_coleta_idStatus_coleta] <> 1 " & _
'		response.write sql
'		response.end					 
	else
		sql = "SELECT " & _
					 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
					 "B.[numero_solicitacao_coleta], " & _
					 "B.[data_solicitacao] " & _
					 "FROM [marketingoki2].[dbo].[solicitacao_coleta_has_pontos_coleta] AS A " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
					 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
					 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
					 " and B.[isMaster] = 1  and left(B.[numero_solicitacao_coleta],1) <> 'E' " & _
					 " order by B.[numero_solicitacao_coleta]" 
'					 " and B.[Status_coleta_idStatus_coleta] <> 3 and B.[Status_coleta_idStatus_coleta] <> 4 and B.[Status_coleta_idStatus_coleta] <> 1 " & _
'		response.write sql
'		response.end			 
	end if				 
	'response.write sql & "<br />"			 
	'response.end
	call search(sql, arr, intarr)
	with response
		If intarr > -1 Then
			For i=0 To intarr
				.Write "<tr>"
				If i Mod 2 = 0 Then
					.Write "<td width='5%' align='center' class='classColorRelPar' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(1,i)&"</td>"
					if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
						.Write 	"<td class='classColorRelPar' align='center' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&DateRight(FormatDateTime(arr(2,i),2))&"</td>"					
					else
						.Write 	"<td class='classColorRelPar' align='center' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&FormatDateTime(arr(2,i),2)&"</td>"					
					end if
					'.Write "<td width='5%' align='center' class='classColorRelPar'>Imprimir</td>"
					.Write	"<td align='center' class='classColorRelPar' style=""cursor:pointer;"" onClick=""javascript:window.open('frmCartaMasterNF.asp?IdSolicitacaoColeta="&arr(0,i)&"&Acao=1&TipoPessoa=','','width=720,height=600,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">Imprimir</td>"
				Else
					.Write "<td width='5%' align='center' class='classColorRelImpar' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(1,i)&"</td>"					
					if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
						.Write 	"<td class='classColorRelImpar' align='center' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&DateRight(FormatDateTime(arr(2,i),2))&"</td>"					
					else
						.Write 	"<td class='classColorRelImpar' align='center' style=""cursor:pointer;"" onClick=""javascript:window.open('frmviewsolpontocoletamaster.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&FormatDateTime(arr(2,i),2)&"</td>"					
					end if
					.Write "<td width='5%' align='center' class='classColorRelImpar'>Imprimir</td>"
				End If
				
				.Write "</tr>"
			Next
		Else
			.Write "<tr>"					
			.Write	"<td colspan='2' align='center' class='classColorRelPar'><b>Nenhum Solicitação Encontrada</b></td>"
			.Write "</tr>"
		End If
	end with
end sub

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
		<form action="" name="frmacompanhasolicitacaoponto" method="POST">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
						<table cellspacing="3" cellpadding="2" width="100%" border=0>
							<tr>
								<td id="explaintitle" align="center">Acompanhamento de Solicitação de Coleta</td>
							</tr>
							<tr>
								<td align="right"><a class="linkOperacional" href="javascript:window.location.href='frmoperacionalponto.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td id="explaintitle" align="right">
									Número da solicitação de Coleta : 
									<INPUT type="text" id="txtNumColeta" name="txtNumColeta" class="text">
									<INPUT type="submit" value="Pesquisar" id="btSearch" name="btSearch" class="btnform">
								</td>
							</tr>
							<tr>
								<td>
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
										<tr>
											<th width=30%>N° Solicitação</th>
											<th width=20%>Dt. Solicitação</th>
											<th width=50%>Carta de Remessa</th>
										</tr>
										<%
										Call getAcompanhaSolicitacao()
										%>
									</table>
								</td>
							</tr>
						</table>
					</div>
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
