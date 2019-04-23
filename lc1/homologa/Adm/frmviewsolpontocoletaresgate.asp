<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
	dim numero_solicitacao
	dim data_resgate
	dim cod_ponto
	dim moeda
	dim documento_baixa
	dim data_baixa

	Sub GetSolicitacao()
		Dim sSql, arr, intarr, i
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
			  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Ponto] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Cadastro_Bonus] AS B " & _
			  "ON A.[cod_bonus] = B.[cod_bonus] " & _
			  "WHERE A.[idSolicitacoes_resgate] = " & Request.QueryString("idsolic")
'		response.write sSql
'		response.end
'		Response.Write Session("IDAdministrator")
'		Response.End()
		Call search(sSql, arr, intarr)
		If intSolicitacao > -1 Then
			For i=0 To intSolicitacao
				numero_solicitacao = arr(6,i)
				data_resgate = arr(5,i)
				cod_ponto = arr(9,i)
				moeda = arr(10,i)
				documento_baixa = arr(3,i)
				data_baixa = arr(4,i)
			Next
		End If
	End Sub

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		if not isnull(sData) and not isempty(sData) then
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
		else
			DateRight = ""
		end if
	End Function

	function getProdutos(idsolicitacao)
		dim sql, arr, intarr, i
		dim html, style
		sql = "SELECT A.[idproduto] " & _
			  ",A.[quantidade] " & _
			  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Ponto] AS A WHERE A.[idSolicitacoes_resgate] = " & idsolicitacao
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				html = html & "<tr>"
				html = html & "<td "&style&">"&arr(0,i)&"</td>"
				html = html & "<td "&style&">"&arr(1,i)&"</td>"
				html = html & "</tr>"
			next
		else
			html = html & "<td "&style&"><b>Nenhum registro encontrado</b></td>"
		end if
		getProdutos = html
	end function

	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia
		Dia = Left(sDate, 2)
		Mes = Mid(sDate, 4, 2)
		Mes = Replace(Mes, "/" ,"")
		If Len(Mes) = 1 Then
			Mes = "0" & Mes
		End If
		Ano = Right(sDate, 4)

		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function

	call GetSolicitacao()
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<style>
	label {
		font-weight:bold;
	}
</style>
<script language="javascript">
function validateForm() {
	var error = 0;

	if (document.frmEditSolicitacaoEntrega.txtQtdCatuchosRecebidos == '') {
		error++;
	}

	if (error == 0) {
		document.frmEditSolicitacaoEntrega.submit();
	}
}
</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="" name="frmEditSolicitacaoEntrega" method="POST">
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Administrar Solicitação de Resgate</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de resgate: </label></td>
				<td><%= numero_solicitacao %> <img src="img/buscar.gif" class="imgexpandeinfo" align="absmiddle" alt="Buscar Solicitações que compuseram a solicitação Resgate" onClick="javascript:window.open('frmviewsolicitacaocompoeresgateponto.asp?idsolic=<%=numero_solicitacao%>','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"/></td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">DT. Solic. Resgate </label></td>
				<td><%= DateRight(data_resgate) %></td>
			</tr>
			<tr id="tridcliente">
				<td width="35%" align="right"><label id="idcliente">Cód. Ponto Coleta </label></td>
				<td><%= cod_ponto %></td>
			</tr>
			<tr id="trrazaosocial">
				<td width="35%" align="right"><label id="razaosocial">Moeda: </label></td>
				<td><%= moeda %></td>
			</tr>
			<tr id="trnomefantasia">
				<td width="35%" align="right"><label id="nomefantasia">Documento Baixa: </label></td>
				<td><%= documento_baixa %></td>
			</tr>
			<tr id="trnomefantasia">
				<td width="35%" align="right"><label id="nomefantasia">Data Baixa: </label></td>
				<td><%= DateRight(data_baixa) %></td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">
					<div style="width:100%;height:155px;overflow:scroll;">
						<table cellpadding="1" cellspacing="1" width="100%" id="tableRelCategories">
							<tr>
								<th>ID. Produto</th>
								<th>Quantidade</th>
							</tr>
							<%= getProdutos(Request.QueryString("idsolic")) %>
						</table>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
