<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	dim idkardex
	dim codigo_cliente
	dim data_recebimento
	dim codigo_produto
	dim descricao_produto
	dim qtd
	dim data_geracao_bonus
	dim numero_solicitacao_coleta
	
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
	
	sub getKardexById(id)
		dim sql, arr, intarr, i
		sql = "SELECT [idKardex] " & _
				  ",[codigo_cliente] " & _
				  ",[data_recebimento] " & _
				  ",[codigo_produto] " & _
				  ",[descricao_produto] " & _
				  ",[qtd] " & _
				  ",[data_geracao_bonus] " & _
				  ",[numero_solicitacao_coleta] " & _
			  "FROM [marketingoki2].[dbo].[Kardex] " & _
			  "WHERE [idKardex] = " & id
'		response.write sql
'		response.end	  
		call search(sql, arr, intarr)			  
		if intarr > -1 then
			for i=0 to intarr
				idkardex = arr(0,i)
				codigo_cliente = arr(1,i)
				data_recebimento = arr(2,i)
				codigo_produto = arr(3,i)
				descricao_produto = arr(4,i)
				qtd = arr(5,i)
				data_geracao_bonus = arr(6,i)
				numero_solicitacao_coleta = arr(7,i)
			next	
		else
			response.write "<script>alert('Kardex não encontrado')</script>"
		end if
	end sub
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		if request.form("action") = "Editar" then
			call editar(request.form("id_kardex"))
		else
			call deletar(request.form("id_kardex"))
		end if
	else
		if request.querystring("id") <> "" then
			call getKardexById(request.querystring("id"))
		end if	
	end if
	
	sub editar(id)
		dim sql
		dim data_recebimento
		dim data_geracao_bonus
		  if request.form("data_receb") <> "" then
		  	data_recebimento = request.form("data_receb")
			data_recebimento = FormatDate(data_recebimento)
			data_recebimento = "convert(datetime, '"&data_recebimento&"')"
		  else
			data_recebimento = "NULL"
		  end if	  
		  if request.form("data_geracao") <> "" then
		  	data_geracao_bonus = request.form("data_geracao")
 			data_geracao_bonus = FormatDate(data_geracao_bonus)
		    data_geracao_bonus = "convert(datetime, '"&data_geracao_bonus&"')"
		  else 
			  data_geracao_bonus = "NULL"		  
		  end if
		sql = "UPDATE [marketingoki2].[dbo].[Kardex] " & _
				   "SET [codigo_cliente] = "&request.form("cod_cliente")&" " & _
					  ",[data_recebimento] = "&data_recebimento&" " & _
					  ",[codigo_produto] = '"&request.form("cod_prod")&"' " & _
					  ",[descricao_produto] = '"&request.form("desc_prod")&"' " & _
					  ",[qtd] = "&request.form("qtd_prod")&" " & _
					  ",[data_geracao_bonus] = "&data_geracao_bonus&" " & _
					  ",[numero_solicitacao_coleta] = '"&request.form("num_solicitacao")&"' " & _
				 "WHERE [idKardex] = " & id
'		response.write sql
'		response.end		 
		call exec(sql)		 
		response.redirect "frmadmkardex.asp"
	end sub
	
	sub deletar(id)
		dim sql
		dim numero
		numero = getNumById(id)
		call updateReturnStatus(numero)
		sql = "delete from kardex where numero_solicitacao_coleta = '" & numero & "'"
		call exec(sql)
		response.redirect "frmadmkardex.asp"
	end sub

	function isMaster(num)
		dim sql, arr, intarr, i
		sql = "select ismaster from solicitacao_coleta where numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to inrarr
				if cint(arr(0,i)) = 0 then
					isMaster = false
				else
					isMaster = true
				end if	
			next
		end if
	end function
	
	sub updateReturnStatus(numeroSol)
		dim sql, arr, intarr, i
		if isMaster(numeroSol) then
			sql = "select b.numero_solicitacao_coleta from solicitacoes_baixadas as a " & _
					"left join solicitacao_coleta as b " & _
					"on a.id_solicitacao = b.idsolicitacao_coleta " & _
					"where a.is_baixada = 1 and a.numero_solicitacao_master = '"&numeroSol&"'"
			call search(sql, arr , intarr)
			if intarr > -1 then
				for i=0 to intarr
					sql = "update solicitacao_coleta set data_recebimento = null, status_coleta_idstatus_coleta = 8, qtd_cartuchos_recebidos = null where numero_solicitacao_coleta = '"&arr(0,i)&"'"
					call exec(sql)
					sql = "delete from solicitacoes_coleta_has_produtos where solicitacao_coleta_idsolicitacoes_coleta = " & getIDSolicitacao(arr(0,i))
					call exec(sql)
				next
			end if		
			sql = "update solicitacao_coleta set data_recebimento = null, status_coleta_idstatus_coleta = 7, qtd_cartuchos_recebidos = null where numero_solicitacao_coleta = '"&numeroSol&"'"
			call exec(sql)
			sql = "delete from solicitacoes_coleta_has_produtos where solicitacao_coleta_idsolicitacoes_coleta = " & getIDSolicitacao(numeroSol)
			call exec(sql)
			sql = "delete from kardex where numero_solicitacao_coleta = '"&numeroSol&"'"
			call exec(sql)
		else
			sql = "update solicitacao_coleta set data_recebimento = null, status_coleta_idstatus_coleta = 7, qtd_cartuchos_recebidos = null where numero_solicitacao_coleta = '"&numeroSol&"'"
			call exec(sql)
			sql = "delete from solicitacoes_coleta_has_produtos where solicitacao_coleta_idsolicitacoes_coleta = " & getIDSolicitacao(numeroSol)
			call exec(sql)
			sql = "delete from kardex where numero_solicitacao_coleta = '"&numeroSol&"'"
			call exec(sql)
		end if
	end sub
	
	function getIDSolicitacao(num)
		dim sql, arr, intarr, i
		sql = "select idsolicitacao_coleta from solicitacao_coleta where numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			getIDSolicitacao = arr(0,0)
		else
			getIDSolicitacao = -1
		end if
	end function
	
	function getNumById(id)
		dim sql, arr, intarr, i
		sql = "select numero_solicitacao_coleta from kardex where idkardex = "&id
		call search(sql, arr, intarr)
		if intarr > -1 then
			getNumById = arr(0,0)
		else
			getNumById = -1
		end if
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

	function validaBonusGerado(num)
		dim sql, arr, intarr, i
		dim contBonus
		contBonus = 0
		sql = "select * from kardex where numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if arr(6,i) <> "" then
					contBonus = contBonus + 1
				end if
			next
			if contBonus > 0 then	
				validaBonusGerado = false
			else
				validaBonusGerado = true
			end if	
		else
			validaBonusGerado = false
		end if
	end function

%>
<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
						<form action="frmeditkardex.asp" name="frmeditkardex" method="POST">
						<input type="hidden" name="id_kardex" value="<%=idkardex%>" />
						<table cellspacing="1" cellpadding="1" width="100%" id="tablelisttransportadoras">
							<tr>
								<td id="explaintitle" colspan="2" align="center">Edição de Kardex</td>
							</tr>
							<tr>
								<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmadmkardex.asp';">&laquo Voltar</a></td>
							</tr>
							<!--tr>
								<td align="right" width="25%">ID:</td>
								<td align="left"><input name="id" type="text" class="textreadonly" size="10" value="<%= idkardex %>" readonly="true" /></td>
							</tr-->
							<tr>
								<td align="right">Cód. Cliente:</td>
								<td align="left"><input name="cod_cliente" type="text" class="text" value="<%= codigo_cliente %>" readonly="true" size="20" /></td>
							</tr>
							<%if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then%>
								<tr>
									<td align="right">Data Receb:</td>
									<td align="left"><input name="data_receb" type="text" class="text" value="<%= DateRight(data_recebimento) %>" readonly="true" size="20" /></td>
								</tr>
							<%else%>
								<tr>
									<td align="right">Data Receb:</td>
									<td align="left"><input name="data_receb" type="text" class="text" value="<%= data_recebimento %>" readonly="true" size="20" /></td>
								</tr>
							<%end if%>
							<tr>
								<td align="right">Cód. Prod:</td>
								<td align="left"><input name="cod_prod" type="text" class="text" value="<%= codigo_produto %>" readonly="true" size="20" /></td>
							</tr>
							<tr>
								<td align="right">Desc. Prod:</td>
								<td align="left"><textarea name="desc_prod" cols="40" rows="5" readonly="true"><%= descricao_produto %></textarea></td>
							</tr>
							<tr>
								<td align="right">Qtd:</td>
								<td align="left"><input name="qtd_prod" type="text" class="text" readonly="true" value="<%= qtd %>" size="20" /></td>
							</tr>
							<%if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then%>
							<tr>
								<td align="right">Data Geração Bônus:</td>
								<td align="left"><input name="data_geracao" type="text" class="textreadonly" value="<%= DateRight(data_geracao_bonus) %>" readonly="true" size="20" /></td>
							</tr>
							<%else%>
							<tr>
								<td align="right">Data Geração Bônus:</td>
								<td align="left"><input name="data_geracao" type="text" class="textreadonly" value="<%= data_geracao_bonus %>" readonly="true" size="20" /></td>
							</tr>
							<%end if%>
							<tr>
								<td align="right">Número Solicitação:</td>
								<td align="left"><input name="num_solicitacao" type="text" class="text" value="<%= numero_solicitacao_coleta %>" readonly="true" size="20" /></td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<%if validaBonusGerado(numero_solicitacao_coleta) then%>
							<tr>
								<td colspan="2" align="center">
									<input name="action" type="submit" class="btnform" value="Deletar" />
								</td>
							</tr>
							<%end if%>	
						</table>
						</form>
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
