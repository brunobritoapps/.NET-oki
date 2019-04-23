<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call getSolicitacoesColeta(request("txtNumColeta"))
		else
			Call getSolicitacoesColeta(0)
		End If
	End Sub	

	Sub getSolicitacoesColeta(sNumSE)
		Dim sSql, arrSolColeta, intSolColeta, i

		if sNumSE = "0" or len(trim(sNumSE)) = 0 then
			sSql = "SELECT " & _
						 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
						 "B.[numero_solicitacao_coleta], " & _
						 "B.[data_solicitacao] " & _
						 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
						 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
						 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
						 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
						 " AND B.[Status_coleta_idStatus_coleta] = 2"
		else
			sSql = "SELECT " & _
						 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
						 "B.[numero_solicitacao_coleta], " & _
						 "B.[data_solicitacao] " & _
						 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
						 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
						 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
						 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
						 " AND B.[Status_coleta_idStatus_coleta] = 2 AND B.[numero_solicitacao_coleta] = '" & sNumSE & "'"
		end if			 

'Response.Write ssql
'Response.End
		
		Call search(sSql, arrSolColeta, intSolColeta)

		With Response	
			If intSolColeta > -1 Then
				For i=0 To intSolColeta
					.Write "<tr style=""cursor:pointer;"" onClick=""javascript:window.open('frmEditSolicitacaoEntrega.asp?idsolic="&arrSolColeta(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"
					If i Mod 2 = 0 Then
						.Write "<td width='5%' align='center' class='classColorRelPar'>"&arrSolColeta(1,i)&"</td>"
						.Write 	"<td class='classColorRelPar' align='center'>"&DateRight(FormatDateTime(arrSolColeta(2,i),2))&"</td>"					
					Else
						.Write "<td width='5%' align='center' class='classColorRelImpar'>"&arrSolColeta(1,i)&"</td>"					
							
						.Write 	"<td class='classColorRelImpar' align='center'>"&DateRight(FormatDateTime(arrSolColeta(2,i),2))&"</td>"
					End If					
					.Write "</tr>"
				Next
			Else
				.Write "<tr>"					
				.Write	"<td colspan='2' align='center' class='classColorRelPar'><b>Nenhum Solicitação Encontrada</b></td>"
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
	
	Sub GetColectType()
		Dim sSql, arrType, intType, i
		Dim sSelected
		
		sSql = "SELECT " & _ 
						"[idCategorias], " & _ 
						"[descricao] " & _
						"FROM [marketingoki2].[dbo].[Categorias] " & _
						"WHERE [ativo] = 1"

		Call search(sSql, arrType, intType)				
		
		If intType > -1 Then
			With Response
				For i=0 To intType
					If Request.QueryString("TypeColect") = CStr(arrType(0,i)) Then
						sSelected = "selected"
					Else
						sSelected = ""
					End If
					.Write "<option value='"&arrType(0,i)&"' "&sSelected&">"&arrType(1,i)&"</option>"															
				Next
			End With
		End If
	End Sub
	
	sub getAcompanhaSolicitacao()
		dim sql, arr, intarr, i
		sql = "SELECT " & _
					 "A.[Solicitacao_coleta_idSolicitacao_coleta], " & _
					 "B.[numero_solicitacao_coleta], " & _
					 "B.[data_solicitacao] " & _
					 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta] AS B  " & _
					 "ON A.[Solicitacao_coleta_idSolicitacao_coleta] = B.[idSolicitacao_coleta]  " & _
					 "WHERE A.[Pontos_coleta_idPontos_coleta] = " & Session("IDPonto") & _
					 " AND B.[Status_coleta_idStatus_coleta] <> 1"
					 
		call search(sql, arr, intarr)
		with response
			If intarr > -1 Then
				For i=0 To intarr
					.Write "<tr style=""cursor:pointer;"" onClick=""javascript:window.open('frmViewSolPontoColeta.asp?idsolic="&arr(0,i)&"','','width=500,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"
					If i Mod 2 = 0 Then
						.Write "<td width='5%' align='center' class='classColorRelPar'>"&arr(1,i)&"</td>"
						.Write 	"<td class='classColorRelPar' align='center'>"&FormatDateTime(arr(2,i),2)&"</td>"					
					Else
						.Write "<td width='5%' align='center' class='classColorRelImpar'>"&arr(1,i)&"</td>"					
							
						.Write 	"<td class='classColorRelImpar' align='center'>"&FormatDateTime(arr(2,i),2)&"</td>"
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
		<form action="" name="frmSearchSE" method="POST">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td colspan="5" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmoperacionalponto.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="2" id="explaintitle" align="center">Solicitações de Coleta para Baixa</td>
							<td id="explaintitle" align="right">
								Número da solicitação de coleta : 
								<INPUT type="text" id=txtNumColeta name=txtNumColeta class="text">
								<INPUT type="submit" value="Pesquisar" id=btSearch name=btSearch class="btnform">
							</td>
						</tr>
						<tr>
							<td colspan="5">
								<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
									<tr>
										<th width=30%>N° Solicitação</th>
										<th width=70%>Dt. Solicitação</th>
									</tr>
									<%
									Call SubmitForm()
									%>
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
