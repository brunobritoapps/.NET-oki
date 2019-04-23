<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub getSolicitacoesColeta()
		Dim sSql, arrSolColeta, intSolColeta, i

		If CInt(Request.QueryString("StatusSol")) = 0 And CInt(Request.QueryString("TypeColect")) = 0 Then
			sSql = "SELECT " & _ 
							"[A].[idSolicitacao_coleta], " & _
							"[A].[numero_solicitacao_coleta], " & _ 
							"[A].[qtd_cartuchos], " & _ 
							"[A].[data_recebimento], " & _ 
							"[A].[motivo_status], " & _ 
							"[A].[isMaster], " & _
							"[C].[nome_fantasia], " & _
							"[C].[typeColect], " & _
							"[B].[typeColect] " & _
							"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _ 
							"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _ 
							"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _ 
							"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
							"ON [B].[Clientes_idClientes] = [C].[idClientes]"
		Else
			If CInt(Request.QueryString("StatusSol")) > 0 And CInt(Request.QueryString("TypeColect")) = 0 Then
				sSql = "SELECT " & _ 
								"[A].[idSolicitacao_coleta], " & _
								"[A].[numero_solicitacao_coleta], " & _ 
								"[A].[qtd_cartuchos], " & _ 
								"[A].[data_recebimento], " & _ 
								"[A].[motivo_status], " & _ 
								"[A].[isMaster], " & _
								"[C].[nome_fantasia], " & _
								"[C].[typeColect], " & _
								"[B].[typeColect] " & _
								"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _ 
								"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
								"ON [B].[Clientes_idClientes] = [C].[idClientes] " & _ 
								"WHERE " & _
								"[A].[Status_coleta_idStatus_coleta] = " & Request.QueryString("StatusSol")
			ElseIf CInt(Request.QueryString("StatusSol")) = 0 And CInt(Request.QueryString("TypeColect")) > 0 Then					
				sSql = "SELECT " & _ 
								"[A].[idSolicitacao_coleta], " & _
								"[A].[numero_solicitacao_coleta], " & _ 
								"[A].[qtd_cartuchos], " & _ 
								"[A].[data_recebimento], " & _ 
								"[A].[motivo_status], " & _ 
								"[A].[isMaster], " & _
								"[C].[nome_fantasia], " & _
								"[C].[typeColect], " & _
								"[B].[typeColect] " & _
								"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _ 
								"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
								"ON [B].[Clientes_idClientes] = [C].[idClientes] " & _ 
								"WHERE " & _
								"[C].[Categorias_idCategorias] = " & Request.QueryString("TypeColect")
			Else
				sSql = "SELECT " & _ 
								"[A].[idSolicitacao_coleta], " & _
								"[A].[numero_solicitacao_coleta], " & _ 
								"[A].[qtd_cartuchos], " & _ 
								"[A].[data_recebimento], " & _ 
								"[A].[motivo_status], " & _ 
								"[A].[isMaster], " & _
								"[C].[nome_fantasia], " & _
								"[C].[typeColect], " & _
								"[B].[typeColect] " & _
								"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS [A] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [B] " & _ 
								"ON [A].[idSolicitacao_coleta] = [B].[Solicitacao_coleta_idSolicitacao_coleta] " & _ 
								"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
								"ON [B].[Clientes_idClientes] = [C].[idClientes] " & _ 
								"WHERE " & _
								"[A].[Status_coleta_idStatus_coleta] = " & Request.QueryString("StatusSol") & _
								" AND [C].[Categorias_idCategorias] = " & Request.QueryString("TypeColect")
			End If							
		End If
		
		Call search(sSql, arrSolColeta, intSolColeta)

		With Response	
			If intSolColeta > -1 Then
				For i=0 To intSolColeta
					If i Mod 2 = 0 Then
						.Write "<tr>"
						.Write "<td width='5%' align='center' class='classColorRelPar'><img class='imgexpandeinfo' src='img/buscar.gif' alt='Verificar Solicitação de Coleta' onClick=""javascript:window.open('frmEditSolicitacaoColetaAdm.asp?iscoletadomiciliar="&arrSolColeta(8,i)&"&idsolic="&arrSolColeta(0,i)&"','','width=500,height=650,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" ></td>"
						.Write 	"<td class='classColorRelPar' align='center'>"&arrSolColeta(6,i)&"</td>"
						.Write 	"<td width='15%' class='classColorRelPar' align='center'>"&arrSolColeta(1,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td align='center' width='15%' class='classColorRelPar'>##/##/####</td>"
						Else
							.Write	"<td class='classColorRelPar'>"&arrSolColeta(3,i)&"</td>"
						End If
						.Write	"<td class='classColorRelPar'  width='15%' align='center'>"&arrSolColeta(2,i)&"</td>"
						If IsNull(arrSolColeta(4,i)) Or arrSolColeta(4,i) = "NULL"  Then
							.Write "<td class='classColorRelPar'  width='15%' align='center'> --- </td>"
						Else
							.Write "<td class='classColorRelPar'  width='15%' align='center'>"&arrSolColeta(4,i)&"</td>"
						End If
						.Write "</tr>"
					Else
						.Write "<tr>"
						.Write "<td width='5%' align='center' class='classColorRelImpar'><img class='imgexpandeinfo' src='img/buscar.gif' alt='Verificar Solicitação de Coleta' onClick=""javascript:window.open('frmEditSolicitacaoColetaAdm.asp?iscoletadomiciliar="&arrSolColeta(8,i)&"&idsolic="&arrSolColeta(0,i)&"','','width=500,height=650,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" ></td>"
						.Write 	"<td class='classColorRelImpar' align='center'>"&arrSolColeta(6,i)&"</td>"
						.Write 	"<td width='15%' class='classColorRelImpar' align='center'>"&arrSolColeta(1,i)&"</td>"
						If IsNull(arrSolColeta(3,i)) Then
							.Write	"<td align='center' width='15%' class='classColorRelImpar'>##/##/####</td>"
						Else
							.Write	"<td class='classColorRelImpar'>"&arrSolColeta(3,i)&"</td>"
						End If
						.Write "<td class='classColorRelImpar'  width='15%' align='center'>"&arrSolColeta(2,i)&"</td>"
						If IsNull(arrSolColeta(4,i)) Or arrSolColeta(4,i) = "NULL" Then
							.Write "<td class='classColorRelImpar'  width='15%' align='center'> --- </td>"
						Else
							.Write "<td class='classColorRelImpar'  width='15%' align='center'>"&arrSolColeta(4,i)&"</td>"
						End If
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
		<form action="frmOperacionalAdm.asp" name="frmOperacionalAdm" method="POST">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td colspan="4" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="2" id="explaintitle" align="center">Solicitações de Coleta</td>
							<td id="explaintitle" align="right">
								Status : 
								<select name="cbStatusColeta" class="select" onChange="window.location.href='frmSolicitacoesAdm.asp?StatusSol=' + this.value + '&TypeColect=' + document.frmOperacionalAdm.cbTipoColeta.value;">
									<option value="0">Todas</option>
									<%Call getStatusColeta()%>
								</select>	
							</td>
							<td id="explaintitle" align="center">
								Categoria : 
								<select name="cbTipoColeta" class="select" onChange="window.location.href='frmSolicitacoesAdm.asp?TypeColect=' + this.value + '&StatusSol=' + document.frmOperacionalAdm.cbStatusColeta.value;">
									<option value="0">Todas</option>
									<%Call GetColectType()%>
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
										<th>DT. Recebimento</th>
										<th>Quantidade</th>
										<th>Motivo Status</th>
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
