<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetSolicitacao()
		Dim sSql, arrSol, intSol, i
		Dim arrPonto
		Dim arrEndColeta

		sSql = "SELECT " & _
				"A.[idSolicitacao_coleta], " & _ 
				"A.[Status_coleta_idStatus_coleta], " & _ 
				"A.[numero_solicitacao_coleta], " & _ 
				"A.[qtd_cartuchos], " & _ 
				"A.[qtd_cartuchos_recebidos], " & _ 
				"A.[data_solicitacao], " & _ 
				"A.[data_aprovacao], " & _ 
				"A.[data_programada], " & _ 
				"A.[data_envio_transportadora], " & _ 
				"A.[data_entrega_pontocoleta], " & _ 
				"A.[data_recebimento], " & _ 
				"A.[motivo_status], " & _			
				"A.[isMaster], " & _ 
				"B.[typeColect], " & _ 
				"B.[Pontos_coleta_idPontos_coleta], " & _ 
				"B.[Contatos_idContatos], " & _ 
				"B.[Clientes_idClientes], " & _ 
				"B.[numero_endereco_coleta], " & _ 
				"B.[comp_endereco_coleta], " & _ 
				"B.[ddd_resp_coleta], " & _ 
				"B.[telefone_resp_coleta], " & _ 
				"B.[contato_coleta], " & _  
				"B.[logradouro_coleta], " & _  
				"B.[bairro_coleta], " & _  
				"B.[municipio_coleta], " & _  
				"B.[estado_coleta], " & _  
				"B.[cep_coleta] " & _  
				"FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A " & _
				"LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B " & _ 
				"ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
				"WHERE A.[idSolicitacao_coleta] = " & Request.QueryString("idsol")
				
		'idsolicitacao					= 0
		'status da solicitacao			= 1
		'numero solicitacao				= 2
		'qtd cartuchos					= 3
		'qtd cartuchos recebidos		= 4
		'data solicitacao				= 5
		'data aprovacao					= 6
		'data programada				= 7
		'data envio para transportadora = 8
		'data entrega ponto de coleta	= 9
		'data recebimento				= 10
		'motivo status					= 11
		'é master						= 12
		'tipo de coleta					= 13
		'id do ponto de coleta			= 14
		'id do contato					= 15
		'id do cliente					= 16
		'numero endereco coleta			= 17
		'comp do endereco coleta		= 18
		'ddd resp coleta				= 19
		'telefone resp coleta			= 20
		'contato coleta					= 21
		'logradouro coleta				= 22
		'bairro coleta					= 23
		'municipio coleta				= 24
		'estado coleta					= 25
		'cep coleta						= 26		
								
		'response.write sSql
		'response.End		
				
		Call search(sSql, arrSol, intSol)		
		If intSol > -1 Then
			Response.Write "<table cellpadding=""1"" cellspacing=""1"" width=""750"" align=""left"" id=""tableRelSolPendente"">"
			Response.Write "<tr>"
			Response.Write "<td width=""40%"" align=""right""><label>Status da Solicitação</label></td>"
			Response.Write "<td>"&GetStatusColeta(arrSol(1,0))&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Número da Solicitação</label></td>"
			Response.Write "<td>"&arrSol(2,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Qtd. Cartuchos</label></td>"
			Response.Write "<td>"&arrSol(3,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Qtd. Cartuchos Recebidos</label></td>"
			Response.Write "<td>"&arrSol(4,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Data Solicitação</label></td>"
			if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
				Response.Write "<td>"&DateRight(FormatDateTime(arrSol(5,0), 2))&"</td>"
			else
				Response.Write "<td>"&FormatDateTime(arrSol(5,0), 2)&"</td>"
			end if	
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Data Aprovação</label></td>"
			if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
				If arrSol(6,0) <> "" Then 
					Response.Write "<td>"&DateRight(FormatDateTime(arrSol(6,0), 2))&"</td>"
				Else
					Response.Write "<td>##/##/####</td>"	
				End If
			else
				If arrSol(6,0) <> "" Then 
					Response.Write "<td>"&FormatDateTime(arrSol(6,0), 2)&"</td>"
				Else
					Response.Write "<td>##/##/####</td>"	
				End If
			end if	
			Response.Write "</tr>"
			If arrSol(13,0) = 1 Then
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Data Envio Transportadora</label></td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					If arrSol(8,0) <> "" Then
						Response.Write "<td>"&DateRight(FormatDateTime(arrSol(8,0), 2))&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End If	
				else
					If arrSol(8,0) <> "" Then
						Response.Write "<td>"&FormatDateTime(arrSol(8,0), 2)&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End If	
				end if	
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Data Programada</label></td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					If arrSol(7,0) <> "" Then 
						Response.Write "<td>"&DateRight(FormatDateTime(arrSol(7,0), 2))&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End If
				else
					If arrSol(7,0) <> "" Then 
						Response.Write "<td>"&FormatDateTime(arrSol(7,0), 2)&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End If
				end if	
				Response.Write "</tr>"
			Else
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Data Entrega Ponto de Coleta</label></td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					If arrSol(9,0) <> "" Then
						Response.Write "<td>"&DateRight(FormatDateTime(arrSol(9,0), 2))&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End if	
				else
					If arrSol(9,0) <> "" Then
						Response.Write "<td>"&FormatDateTime(arrSol(9,0), 2)&"</td>"
					Else
						Response.Write "<td>##/##/####</td>"
					End if	
				end if	
				Response.Write "</tr>"
			End If
			Response.Write "<tr>"
			Response.Write "<td align=""right""><label>Data Recebimento</label></td>"	
			if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
				If arrSol(10,0) <> "" Then
					Response.Write "<td>"&DateRight(FormatDateTime(arrSol(10,0), 2))&"</td>"
				Else
					Response.Write "<td>##/##/####</td>"
				End If	
			else
				If arrSol(10,0) <> "" Then
					Response.Write "<td>"&FormatDateTime(arrSol(10,0), 2)&"</td>"
				Else
					Response.Write "<td>##/##/####</td>"
				End If	
			end if	
			Response.Write "</tr>"
			If arrSol(1,0) = 3 Or arrSol(1,0) = 4 Then
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Motivo do Status</label></td>"
				Response.Write "<td>"&arrSol(11,0)&"</td>"
				Response.Write "</tr>"
			End If	
			If arrSol(13,0) = 0 Then
				Response.Write "<tr>"
				Response.Write "<td align=""center"" colspan=""2""><label>Ponto de Coleta</label></td>"
				Response.Write "</tr>"
				arrPonto = Split(GetPontoColetaDesc(arrSol(14,0)),";")
				Response.Write "<tr>"
				Response.Write "<td width=""700"" colspan=""2"" align=""center"">"
				Response.Write "<table cellspacing=""1"" cellpadding=""1"" align=""left"" width=""715"" id=""tableRelSolPendente"">"
				Response.Write "<tr>"
				Response.Write "<th>Nome</th>"
				Response.Write "<th>Cep</th>"
				Response.Write "<th>Logradouro</th>"
				Response.Write "<th>Telefone</th>"
				Response.Write "<th>Município</th>"
				Response.Write "<th>Estado</th>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td class='classColorRelPar'>"&arrPonto(0)&"</td>"
				Response.Write "<td class='classColorRelPar'>"&arrPonto(1)&"</td>"
				Response.Write "<td class='classColorRelPar'>"&arrPonto(2)&"/ n° " & arrPonto(4) & " / " & arrPonto(3) & " / " & arrPonto(7) & "</td>"
				Response.Write "<td class='classColorRelPar'> ("& arrPonto(5) & ") - " & arrPonto(6)&"</td>"
				Response.Write "<td class='classColorRelPar'>"&arrPonto(8)&"</td>"
				Response.Write "<td class='classColorRelPar'>"&arrPonto(9)&"</td>"
				Response.Write "</tr>"
				Response.Write "</table>"
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
			Response.Write "<tr>"	
			Response.Write "<td align=""right""><label>Contato que fez a Solicitação de Coleta</label></td>"
			Response.Write "<td>"&GetContatoColeta(arrSol(15,0))&"</td>"
			Response.Write "</tr>"	
			If arrSol(13,0) = 1 Then
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Cep da Coleta</label></td>"
				Response.Write "<td>"&arrSol(26,0)&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Logradouro da Coleta</label></td>"
				Response.Write "<td>"&arrSol(22,0)&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Complemento do End. Coleta</label></td>"
				Response.Write "<td>"&arrSol(19,0)&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Bairro da Coleta</label></td>"
				Response.Write "<td>"&arrSol(23,0)&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Município da Coleta</label></td>"
				Response.Write "<td>"&arrSol(24,0)&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td align=""right""><label>Estado da Coleta</label></td>"
				Response.Write "<td>"&arrSol(25,0)&"</td>"
				Response.Write "</tr>"
			End If
			Response.Write "</table>"
		End If
	End Sub

	Sub GetProductBySol()
		Dim sSql, arrProd, intProd, i
		sSql = "SELECT " & _
						"A.[Produtos_idProdutos], " & _ 
						"A.[quantidade], " & _
						"B.[descricao] " & _ 
						"FROM [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
						"ON B.[IDOki] = A.[Produtos_idProdutos] " & _
						"WHERE A.[Solicitacao_coleta_idSolicitacoes_coleta] = " & Request.QueryString("idsol")
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(0,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(0,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				End If	
			Next
		Else
			Response.Write "<tr><td colspan=""3"" class='classColorRelPar' align=""center""><b>Material ainda não recebido</b></td></tr>"	
		End If				
	End Sub
	
	Function GetStatusColeta(IDStatus)
		Dim sSql, arrStatus, intStatus
		sSql = "select * from Status_coleta where idstatus_coleta = " & IDStatus
		Call search(sSql, arrStatus, intStatus)
		If intStatus > -1 Then
			GetStatusColeta = arrStatus(1,0)	
		End If
	End Function
	
	Function GetContatoColeta(IDContato)
		Dim sSql, arrCon, intCon
		sSql = "select nome from contatos where idcontatos = " & IDContato
		Call search(sSql, arrCon, intCon)
		If intCon > -1 Then
			GetContatoColeta = arrCon(0,0)	
		End If
	End Function
	
	Function GetNomePontoColeta(IDPonto)
		Dim sSql, arrPonto, intPonto
		sSql = "select nome_fantasia from pontos_coleta where idpontos_coleta = " & IDPonto
		Call search(sSql ,arrPonto, intPonto)
		If intPonto > -1 Then
			GetNomePontoColeta = arrPonto(0,0)	
		End If 
	End Function
	
'	Function GetEnderecoColeta(IDEndereco)
'		Dim sSql, arrEnd, intEnd
'		sSql = "select * from cep_consulta where idcep_consulta = " & IDEndereco
'		Call search(sSql, arrEnd, intEnd)
'		If intEnd > -1 Then
'			GetEnderecoColeta = arrEnd(1,0) & ";" & arrEnd(2,0) & ";" & arrEnd(3,0) & ";" & arrEnd(4,0) & ";" & arrEnd(5,0)	& ";" & arrEnd(6,0)
'		End If
'	End Function

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
	
	Function GetPontoColetaDesc(IDPonto)
		Dim sSql, arrPonto, intPonto, i
		Dim Ret

		sSql = "select " & _
				"a.nome_fantasia, " & _
				"a.complemento_endereco, " & _
				"a.numero_endereco, " & _
				"a.ddd, " & _
				"a.telefone, " & _ 
				"a.cep, " & _
				"a.logradouro, " & _
				"a.bairro, " & _
				"a.municipio, " & _
				"a.estado " & _
				"from pontos_coleta as a " & _
				"where a.status_pontocoleta = 1 and a.idpontos_coleta = " & IDPonto

		Call search(sSql, arrPonto, intPonto)
		If intPonto > -1 Then
			Ret = arrPonto(0,0) & ";" & arrPonto(5,0) & ";" & arrPonto(6,0) & ";" & arrPonto(1,0) & ";" & arrPonto(2,0) & ";" & arrPonto(3,0) & ";" & arrPonto(4,0) & ";" & arrPonto(7,0) & ";" & arrPonto(8,0) & ";" & arrPonto(9,0)
		End If
		GetPontoColetaDesc = Ret
	End Function
%>
<html>
<head>
<style>
	label {
		font-weight:bold;
		padding:5px 5px 5px 5px;
	}
</style>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" align="left" id="tableRelSolPendente">
						<tr>
							<td width="750"><%Call GetSolicitacao()%></td>
						</tr>
						<tr>
							<td>
								<div style="width:715px;height:126px;overflow:auto;">
									<table cellpadding="1" cellspacing="1" width="100%" align="left" id="tableRelSolPendente">
										<tr>
											<td colspan="5" id="explaintitle" align="center">Acompanhamento de Solicitação de Coleta</td>
										</tr>
										<tr>
											<th>Cód. Produto</th>
											<th>Descrição Produto</th>
											<th>Quantidade</th>
										</tr>
										<%Call GetProductBySol()%>
									</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>
