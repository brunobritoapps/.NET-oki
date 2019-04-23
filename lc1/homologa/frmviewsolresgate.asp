<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetSolicitacao()
		Dim sSql, arrSol, intSol, i
		Dim arrPonto
		Dim arrEndColeta

		sSql = "SELECT A.[idSolicitacoes_resgate] " & _
					  ",A.[cod_bonus] " & _
					  ",A.[idsolicitacao] " & _
					  ",A.[documento_baixa] " & _
					  ",A.[data_baixa] " & _
					  ",A.[data_solicitacao_resgate] " & _
					  ",A.[numero_solicitacao_geracao] " & _
					  ",A.[idproduto] " & _
					  ",A.[quantidade] " & _
					  ",A.[idcliente] " & _
					  ",B.[moeda] " & _					  
				  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Cadastro_Bonus] AS B " & _
				  "ON A.[cod_bonus] = B.[cod_bonus] WHERE A.idsolicitacao = " & Request.QueryString("idsol")
				
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
			Response.Write "<td align=""left"" width=""20%""><label>Número da Solicitação</label></td>"
			Response.Write "<td align=""left"">"&arrSol(6,0)&" <img src=""img/buscar.gif"" class=""imgexpandeinfo"" align=""absmiddle"" alt=""Buscar Solicitações que compuseram a solicitação Resgate"" onClick=""javascript:window.open('adm/frmviewsolicitacaocompoeresgate.asp?idsolic="&arrSol(6,0)&"','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');""/></td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""left""><label>DT. Solic. Resgate</label></td>"
			if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
				Response.Write "<td align=""left"">"&DateRight(FormatDateTime(arrSol(5,0), 2))&"</td>"
			else
				Response.Write "<td align=""left"">"&FormatDateTime(arrSol(5,0), 2)&"</td>"
			end if	
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""left""><label>Cód. Cliente</label></td>"
			Response.Write "<td align=""left"">"&arrSol(9,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""left""><label>Moeda</label></td>"
			Response.Write "<td align=""left"">"&arrSol(10,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""left""><label>Documento Baixa</label></td>"
			Response.Write "<td align=""left"">"&arrSol(3,0)&"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""left""><label>Data da Baixa</label></td>"
			if not isnull(arrSol(4,0)) and not isempty(arrSol(4,0)) then
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					Response.Write "<td align=""left"">"&DateRight(FormatDateTime(arrSol(4,0), 2))&"</td>"
				else
					Response.Write "<td align=""left"">"&FormatDateTime(arrSol(4,0), 2)&"</td>"
				end if	
			else	
				Response.Write "<td align=""left""></td>"
			end if	
			Response.Write "</tr>"
			Response.Write "</table>"
		End If
	End Sub

	Sub GetProductBySol()
		Dim sSql, arrProd, intProd, i
		sSql = "SELECT A.[idproduto] " & _
					  ",A.[quantidade] " & _
					  ",B.[descricao] " & _
				  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B ON A.[idproduto] = B.[idoki] " & _
				  "WHERE A.idsolicitacao = " & Request.QueryString("idsol")
		'response.write sSql
'		response.end		  
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
			Response.Write "<tr><td colspan=""3"" class='classColorRelPar' align=""center""><b>Nenhum registro encontrado</b></td></tr>"	
		End If				
	End Sub
	
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
					<table cellpadding="1" cellspacing="1" width="702" align="left" id="tableRelSolPendente">
						<tr>
							<td width="750"><%Call GetSolicitacao()%></td>
						</tr>
						<tr>
							<td>
								<div style="width:750px;height:126px;overflow:auto;">
									<table cellpadding="1" cellspacing="1" width="750" align="left" id="tableRelSolPendente">
										<tr>
											
                      <td colspan="5" id="explaintitle" align="center">Acompanhamento 
                        de Solicitação de Resgate</td>
										</tr>
										<tr>
											<th>Cód. Produto</th>
											<th>Descrição</th>
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
