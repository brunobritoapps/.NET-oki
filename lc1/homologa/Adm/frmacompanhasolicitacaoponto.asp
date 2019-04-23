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
					 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
					 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
					 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
					 " AND B.[Status_coleta_idStatus_coleta] <> 1 and B.[Status_coleta_idStatus_coleta] <> 3 and B.[Status_coleta_idStatus_coleta] <> 4 " & _
					 " and B.[numero_solicitacao_coleta] = '" & numcoleta & "' " & _
					 " order by B.[numero_solicitacao_coleta]"
'		response.write sql
'		response.end					 
	else
		sql = "SELECT " & _
					 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
					 "B.[numero_solicitacao_coleta], " & _
					 "B.[data_solicitacao] " & _
					 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
					 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
					 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
					 " AND B.[Status_coleta_idStatus_coleta] <> 1 and B.[Status_coleta_idStatus_coleta] <> 3 and B.[Status_coleta_idStatus_coleta] <> 4 order by B.[numero_solicitacao_coleta]" 
	end if				 
				 
	call search(sql, arr, intarr)
	with response
		If intarr > -1 Then
			For i=0 To intarr
				.Write "<tr style=""cursor:pointer;"" onClick=""javascript:window.open('frmViewSolPontoColeta.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"
				If i Mod 2 = 0 Then
					.Write "<td width='5%' align='center' class='classColorRelPar'>"&arr(1,i)&"</td>"
					if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
						.Write 	"<td class='classColorRelPar' align='center'>"&DateRight(FormatDateTime(arr(2,i),2))&"</td>"					
					else
						.Write 	"<td class='classColorRelPar' align='center'>"&FormatDateTime(arr(2,i),2)&"</td>"					
					end if
				Else
					.Write "<td width='5%' align='center' class='classColorRelImpar'>"&arr(1,i)&"</td>"					
					if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
						.Write 	"<td class='classColorRelImpar' align='center'>"&DateRight(FormatDateTime(arr(2,i),2))&"</td>"					
					else
						.Write 	"<td class='classColorRelImpar' align='center'>"&FormatDateTime(arr(2,i),2)&"</td>"					
					end if
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

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano
		if sData <> "" then
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
'			response.write Mes & "/" & Dia & "/" & Ano & "<br />"
			DateRight = Mes & "/" & Dia & "/" & Ano
		else
			DateRight = ""
		end if	
	End Function

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
		<form action="frmacompanhasolicitacaoponto.asp" name="frmacompanhasolicitacaoponto" method="POST">
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
									Número da solicitação de coleta : 
									<INPUT type="text" id="txtNumColeta" name="txtNumColeta" class="text">
									<INPUT type="submit" value="Pesquisar" id="btSearch" name="btSearch" class="btnform">
								</td>
							</tr>
							<tr>
								<td>
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
										<tr>
											<th width=30%>N° Solicitação</th>
											<th width=70%>Dt. Solicitação</th>
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
