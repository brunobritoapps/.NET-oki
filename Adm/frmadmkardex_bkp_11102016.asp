<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%

	dim logger
	logger = logger & ""

	function getListKardex()
		dim sql, arr, intarr, i
		dim ret
		dim style

		ret = ""
		style = "class=""classColorRelPar"""

		sql = "SELECT " & _
				"[idKardex], " & _
				"[codigo_cliente], " & _
				"[data_recebimento], " & _
				"[codigo_produto], " & _
				"[descricao_produto], " & _
				"[qtd], " & _
				"[data_geracao_bonus], " & _
				"[numero_solicitacao_coleta] " & _
				"FROM [marketingoki2].[dbo].[Kardex]"
		if trim(request.QueryString("busca")) <> "" then
			sql = sql & " WHERE [numero_solicitacao_coleta] = '"&trim(request.QueryString("busca"))&"'"
		end if
		
		'Response.Write sql	
		'Responde.End 
		
		call search(sql, arr, intarr)
		
		Dim iNumPags
		
		if intarr > -1 then
			ret = ret & "<input type=""hidden"" name=""intsol"" value="""&intarr&""" />"
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = intarr+1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
					
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			ret = ret & "<tr><td colspan=9><div id=pag>"
			ret = ret & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			ret = ret & "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if

				ret = ret & "<tr>"
				ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(7,i)&"</td>"
				ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(1,i)&"</td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&DateRight(arr(2,i))&"</td>"
				else
					ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(2,i)&"</td>"
				end if
				ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(3,i)&"</td>"
				ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(4,i)&"</td>"
				ret = ret & "<td align=""right"" onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(5,i)&"</td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&DateRight(arr(6,i))&"</td>"
				else
					ret = ret & "<td onclick=""window.location.href='frmeditkardex.asp?id="&arr(0,i)&"'"" "&style&" style=""cursor:pointer;"" title=""Administra��o Kardex ["&arr(7,i)&"]"">"&arr(6,i)&"</td>"
				end if
				ret = ret & "</tr>"
			next
			ret = ret & "<tr><td colspan=9><div id=pag>"
			ret = ret & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			ret = ret & "</div></td></tr>"
		else
			ret = ret & "<tr><td align=""center"" colspan=""8"" "&style&"><b>Nenhum registro encontrado</b></td></tr>"
		end if
				
		getListKardex = ret
	end function

	sub deleteSelectedSol()
		dim solicitacoes, i

		solicitacoes = split(request.form("solicitacoes"), ",")
		if ubound(solicitacoes) > -1 then
			for i=0 to ubound(solicitacoes)
				call updateReturnStatus(solicitacoes(i))
			next
		end if
	end sub

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

	function isMasterString(num)
		if left(num, 1) = "M" then
			isMasterString = true
		else
			isMasterString = false
		end if
	end function

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

	' ID do ponto de coleta pela Solicitacao
	function getIDPontoByNumSol(num)
		dim sql, arr, intarr, i
		sql = "select a.pontos_coleta_idpontos_coleta from solicitacao_coleta_has_pontos_coleta as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.solicitacao_coleta_idsolicitacao_coleta = b.idsolicitacao_coleta " & _
				"where b.numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			getIDPontoByNumSol = arr(0,0)
		else
			getIDPontoByNumSol = -1
		end if
	end function

	' ID do cliente pela Solicitacao
	function getIDCliByNumSol(num)
		dim sql, arr, intarr, i
		sql = "select c.idclientes, c.cod_cli_consolidador from solicitacao_coleta as a " & _
				"left join solicitacao_coleta_has_clientes as b " & _
				"on a.idsolicitacao_coleta = b.solicitacao_coleta_idsolicitacao_coleta " & _
				"left join clientes as c " & _
				"on b.clientes_idclientes = c.idclientes " & _
				"where a.numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if arr(1,i) = 0 or arr(1,i) = "" then
					getIDCliByNumSol = arr(0,i)
				else
					getIDCliByNumSol = arr(1,i)
				end if
			next
		else
			getIDCliByNumSol = 0
		end if
	end function

	'====================================================================================================================
	' Bonus / Rotinas referentes ao processamento de B�nus
	'====================================================================================================================
	sub geraBonusByCliente()
		dim solicitacoes, i
		dim idcliente
		dim codbonus
		dim dataGeracao
		dim dataValidade
		dim dataResgate
		dim moeda
		dim validade
		dim descricao
		dim num_solicitacao
		dim cod_produto
		dim pontuacao
		dim pontuacaoTarget
		dim qtd
		dim qtd_prod_sol
		dim pontuacao_calculada
		dim cont_produto
		dim sql, arr, intarr, dataValidade_, mes_,dia_,ano_

		sql = "SELECT " & _
				"[idKardex], " & _
				"[codigo_cliente], " & _
				"[data_recebimento], " & _
				"[codigo_produto], " & _
				"[descricao_produto], " & _
				"[qtd], " & _
				"[data_geracao_bonus], " & _
				"[numero_solicitacao_coleta] " & _
				"FROM [marketingoki2].[dbo].[Kardex] " & _
				"WHERE isnull(data_geracao_bonus, '') = '' "
				if trim(request.QueryString("busca")) <> "" then
					sql = sql & " AND [numero_solicitacao_coleta] = '"&trim(request.QueryString("busca"))&"'"
				end if
		
		'response.write sql
		'response.end
		
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				num_solicitacao = arr(7,i)
				num_solicitacao = trim(num_solicitacao)
				'call updateReturnStatus(solicitacoes(i))
				
				if isMaster(num_solicitacao) then
					idcliente = getIDPontoByNumSol(num_solicitacao)
					codbonus = getBonusByPontoColeta(idcliente)
				else
					idcliente = getIDCliByNumSol(num_solicitacao)
					codbonus = getBonusByCliente(idcliente)
				end if
				
				'Ricardo Silva 03/10/2013
				If codbonus = "" and i < 1 then
				 	Response.Write "<script>alert('B�nus n�o cadastrado para este cliente! N�o ser� gerado b�nus');</script>"
				 'response.end
				End If
				
				dataGeracao = Year(Now()) & "/" & Month(Now()) & "/" & Day(Now())
				
				call getInfoByBonus(codbonus, validade, moeda, descricao)

				'response.write validade
				'response.end
				
				dataValidade = dateadd("d",validade,dataGeracao)
				dataResgate = ""
				
				dataValidade_ = split(dataValidade,"/")
				
				 dia_ = dataValidade_(0)
				 mes_ = dataValidade_(1)
				 ano_ = dataValidade_(2)
				
					if len(dia_) = 1 then 'dia
						dia_ = "0"&dia_
					end if
					
					if len(mes_) = 1 then 'mes
						mes_ = "0"&mes_
					end if
				
				dataValidade =  dia_ & "/" & mes_& "/" & ano_
				
				
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					dataGeracao = DateRight(formatdatetime(dataGeracao, 2))
					'dataValidade = DateRight(dataValidade)
				else
					dataGeracao = formatdatetime(dataGeracao, 2)
					'dataValidade = formatdatetime(dataValidade, 2)
				end if
				
				'response.write " entrou no else do insere bonus gerado" &   FormatDate(dataValidade) 
				'response.end
				
				cod_produto = getProdutoBySolicitacao(num_solicitacao)
				
				for cont_produto=0 to ubound(cod_produto, 2)
					if len(trim(cod_produto(0,cont_produto))) > -1 then
						qtd_prod_sol = cod_produto(1,cont_produto)
						
						call getPontuacaoByProduto(cod_produto(0,cont_produto), codbonus, pontuacao, pontuacaoTarget, qtd)
						'response.write "numero solicitacao:  " & cod_produto(0,cont_produto)& codbonus & num_solicitacao & " " & pontuacao & " " & pontuacaoTarget & " " & cint(qtd) & "<br />"
						'response.end
						if clng(pontuacao) > -1 and clng(pontuacaoTarget) > -1 and clng(qtd) > -1 then
							pontuacao_calculada = (clng(qtd_prod_sol) / clng(qtd)) * pontuacao
							'response.write "insereTabelaTemporaria"
							'response.write ubound(cod_produto, 2)
							'response.write "Total de registros " & intarr
							'response.end
							call insereTabelaTemporaria(idcliente, codbonus, dataGeracao, dataValidade, dataResgate, moeda, descricao, pontuacao_calculada, pontuacaoTarget, num_solicitacao, cod_produto(0,cont_produto))
						    
						else
							call atualizaDataGeracaoBonus(num_solicitacao)
							'response.write  "atualizaDataGeracaoBonus"
							'response.end
						end if
					end if
				next
			next
			'response.write " Vai chamar insereBonusGerado"
			'response.end
			call insereBonusGerado()
		end if
	end sub

	sub atualizaDataGeracaoBonus(numero_solicitacao)
		dim sql
		dim data
		dim arr, intarr
		data = year(now()) & "-" & month(now()) & "-" & day(now())
		sql = "select data_geracao_bonus from kardex where numero_solicitacao_coleta = '"&numero_solicitacao&"'"
		
		'response.write "entrou Linha 325" & data
		'response.end   
		call search(sql ,arr, intarr)
		if not intarr > -1 then
		response.write "entrou Linha 325" & data
		response.end  
			sql = "update kardex set data_geracao_bonus = convert(datetime, '"&data&"') where numero_solicitacao_coleta = '"&numero_solicitacao&"'"
		'	response.write "Passou no update"
		'	response.write sql
		'	response.end
		response.write "entrou Linha 325" & data
		response.end   
			call exec(sql)
		end if
	end sub

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		if sData <> "" then
			'Dia = Left(sData, 2)
			'Dia = Replace(Dia, "/", "")
			Dia = day(sData)
			If Len(Dia) = 1 Then
				Dia = "0" & Dia
			End If

			'If Len(Replace(Left(sData, 2), "/", "")) = 1 Then
			'	Mes = Mid(sData, 3, 2)
			'	Mes = Replace(Mes, "/", "")
			'	If Len(Mes) = 1 Then
			'		Mes = "0" & Mes
			'	End If
			'Else
			'	Mes = Mid(sData, 4, 2)
			'	Mes = Replace(Mes, "/", "")
			'	If Len(Mes) = 1 Then
			'		Mes = "0" & Mes
			'	End If
			'End If

			Mes = month(sData)
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If

			'Ano = Right(sData, 4)
			'Ano = Replace(Ano, "/", "")
			Ano = year(sData)
			If Len(Ano) = 1 Then
				Ano = "0" & Ano
			End If
'			response.write Mes & "/" & Dia & "/" & Ano & "<br />"
			'DateRight = Mes & "/" & Dia & "/" & Ano
			DateRight = Dia & "/" & Mes & "/" & Ano
		else
			DateRight = ""
		end if
	End Function

	function validaSolicitacao(num_solicitacao, codprod)
		dim sql, arr, intarr, i
		dim boolValida
		boolValida = 0
		sql = "select data_recebimento from kardex where numero_solicitacao_coleta = '"&num_solicitacao&"' and codigo_produto = '"&codprod&"'"
'		response.write sql & "<br />"
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if len(arr(0,i)) <> "" and len(arr(0,i)) > 0 then
					boolValida = boolValida + 1
				end if
			next
'			response.write "boolValida: " & boolValida & "<br />"
			if boolValida = 0 then
				validaSolicitacao = true
			else
				validaSolicitacao = false
			end if
		else
			validaSolicitacao = false
		end if
	end function

	function validaDataRecebimento(numero_solicitacao, codbonus)
		dim sql, arr, intarr, i
		dim valida
		sql = "select " & _
				"dbo.fn_dateformat(a.data_recebimento, 1), dbo.fn_dateformat(b.data_inicio_contabilizacao,1) " & _
				"from kardex as a, cadastro_bonus as b " & _
				"where a.numero_solicitacao_coleta = '"&numero_solicitacao&"' and b.cod_bonus = '"&codbonus&"'"
'		response.write sql & "<br />"
'		response.end
		call search(sql, arr, intarr)
'		for each valor in arr
'			response.write "valor: " & valor & "<br />"
'		next
'		response.end
		if intarr > -1 then
			for i=0 to intarr
'				response.write "valida: "&valida&" data1: "&formataValidacao(arr(1,i))&" data2: "&formataValidacao(arr(0,i))&";<br />"
'				response.write datediff("d", arr(1,i), arr(0,i)) & "<br />"
				valida = datediff("d", formataValidacao(arr(1,i)), formataValidacao(arr(0,i)))
			next
'			response.write valida & "<br />"
'			response.end
			if valida => 0 then
				validaDataRecebimento = true
			else
				validaDataRecebimento = false
			end if
		else
			validaDataRecebimento = false
		end if
	end function

	function formataValidacao(sdata)
		dim data
		dim dia
		dim mes
		dim ano
		data = split(sdata,"/")
		dia = data(0)
		if len(dia) = 1 then
			dia = "0"&dia
		end if
		mes = data(1)
		if len(mes) = 1 then
			mes = "0"&mes
		end if
		ano = data(2)
		formataValidacao = mes&"/"&dia&"/"&ano
	end function

	sub insereBonusGerado()
		dim sql, arr, intarr, i
		dim sql_insert
'		response.write "insereBonusGerado"
'		response.end
		sql = "select numero_solicitacao " & _
				", idproduto " & _
				", pontuacao " & _
				", pontuacao_atingir " & _
				", moeda " & _
				", idcliente " & _
				", data_geracao " & _
				", data_validade " & _
				", data_resgate " & _
				", idbonus " & _
				"from ##valida_gerabonus"
				' numero solicitacao = 0
				' idproduto 		= 1
				' pontuacao 		= 2
				' pontuacao atingir = 3
				' moeda 			= 4
				' idcliente 		= 5
				' data geracao 		= 6
				' data validade 	= 7
				' data resgate 		= 8
				' id bonus 			= 9
		'response.write sql
		'response.end
		call search(sql, arr, intarr)
		dim cont
'		for cont=0 to ubound(arr,2)
'			response.write arr(0,cont) & "<br />"
'		next
'		response.end
		if intarr > -1 then
			for i=0 to intarr
				if arr(5,i) > -1 then
'					response.write "arr(5,i): passou aqui: " & arr(5,i) & "<br />"
'					if validaSolicitacao(arr(0,i), arr(1,i)) then
'						response.write "validaSolicitacao(arr(0,i), arr(1,i)): " & validaSolicitacao(arr(0,i), arr(1,i)) & "<br />"
						if validaDataRecebimento(arr(0,i), arr(9,i)) then
'							response.write "validaDataRecebimento(arr(0,i), arr(9,i)): " & validaDataRecebimento(arr(0,i), arr(9,i)) & "<br />"
							if validaInsertBonusPonto(arr(0,i), arr(9,i), arr(1,i)) then
'								response.write "validaInsertBonusPonto(arr(0,i), arr(9,i), arr(1,i)): " & validaInsertBonusPonto(arr(0,i), arr(9,i), arr(1,i)) & "<br />"
								if isMasterString(arr(0,i)) then
									sql_insert = "INSERT INTO [marketingoki2].[dbo].[Bonus_Gerado_PontoColeta] " & _
													   "([Pontos_coleta_idPontos_coleta] " & _
													   ",[cod_bonus] " & _
													   ",[data_geracao] " & _
													   ",[data_validade] " & _
													   ",[data_resgate] " & _
													   ",[moeda] " & _
													   ",[descricao] " & _
													   ",[pontuacao] " & _
													   ",[pontuacao_atingir] " & _
													   ",[idproduto] " & _
													   ",[numero_solicitacao]) " & _
												 "VALUES " & _
													   "("&arr(5,i)&" " & _
													   ",'"&arr(9,i)&"' " & _
													   ",convert(datetime, '"&FormatDate(arr(6,i))&"') " & _
													   ",convert(datetime, '"&FormatDate(arr(7,i))&"') " & _
													   ",NULL " & _
													   ",'"&arr(4,i)&"' " & _
													   ",'"&getDescProduto(arr(1,i))&"' " & _
													   ","&arr(2,i)&" " & _
													   ","&arr(3,i)&" " & _
													   ",'"&arr(1,i)&"' " & _
													   ",'"&arr(0,i)&"')"
								else
								
								'response.write " entrou no else do insere bonus gerado" & FormatDate(arr(6,i)) & FormatDate(arr(7,i))
								'response.end
								
									sql_insert = "INSERT INTO [marketingoki2].[dbo].[Bonus_Gerado_Clientes] " & _
													   "([Clientes_idClientes] " & _
													   ",[cod_bonus] " & _
													   ",[data_geracao] " & _
													   ",[data_validade] " & _
													   ",[data_resgate] " & _
													   ",[moeda] " & _
													   ",[descricao] " & _
													   ",[pontuacao] " & _
													   ",[pontuacao_atingir] " & _
													   ",[idproduto] " & _
													   ",[numero_solicitacao]) " & _
												 "VALUES " & _
													   "("&arr(5,i)&" " & _
													   ",'"&arr(9,i)&"' " & _
													   ",convert(datetime, '"&FormataDataPonto(arr(6,i))&"') " & _
													   ",convert(datetime, '"&FormatDate(arr(7,i))&"') " & _
													   ",NULL " & _
													   ",'"&arr(4,i)&"' " & _
													   ",'"&getDescProduto(arr(1,i))&"' " & _
													   ","&arr(2,i)&" " & _
													   ","&arr(3,i)&" " & _
													   ",'"&arr(1,i)&"' " & _
													   ",'"&arr(0,i)&"')"
								end if
								'response.write "insere bonus: " & sql_insert & "<br />"
								'response.end
								call exec(sql_insert)
							end if
						else
							logger = logger & "<b style=""color:#FF0000;"">Erro: Data Recebimento � menor que a Data de Inicio da Contabiliza��o ["&arr(0,i)&"]</b><br />"
						end if
'					end if
				else
					logger = logger & "<b style=""color:#FF0000;"">Erro: Cliente n�o participa do Programa de B�nus ["&arr(0,i)&"]</b><br />"
				end if
			next
			'reponse.write "executou insere bonus gerado"
			'response.end
			call atualizaKardex()
			call atualizaSaldoBonusCliente()
			call atualizaSaldoBonusPonto()
'			call atualizaBonusPontoColeta()
		else
		end if
	end sub
	
	'Ricardo Silva 
	function formataDataPonto(sdata)
		dim data
		dim dia
		dim mes
		dim ano
		data = split(sdata, "/")
		dia = data(0)
		mes = data(1)
		ano = data(2)
		if len(dia) = 1 then
			dia = "0"&dia
		end if
		if len(mes) = 1 then
			mes = "0"&mes
		end if
		if len(ano) = 1 then
			ano = "0"&ano
		end if
		'formataDataPonto = ano&"/"&mes&"/"&dia
		'formataDataPonto = dia&"/"&mes&"/"&ano  
		formataDataPonto = ano&"-"&mes&"-"&dia  
	end function

	function validaInsertBonusPonto(numero_solicitacao, cod_bonus, idproduto)
		dim sql, arr, intarr, i
		if isMasterString(numero_solicitacao) then
			sql = "SELECT [idBonus_Gerado_PontoColeta] " & _
					  ",[Pontos_coleta_idPontos_coleta] " & _
					  ",[cod_bonus] " & _
					  ",[data_geracao] " & _
					  ",[data_validade] " & _
					  ",[data_resgate] " & _
					  ",[moeda] " & _
					  ",[descricao] " & _
					  ",[pontuacao] " & _
					  ",[pontuacao_atingir] " & _
					  ",[idproduto] " & _
					  ",[numero_solicitacao] " & _
					  ",[saldo] " & _
				  "FROM [marketingoki2].[dbo].[Bonus_Gerado_PontoColeta] " & _
				  "WHERE [numero_solicitacao] = '"&numero_solicitacao&"' AND " & _
				  "[idproduto] = '"&idproduto&"' AND " & _
				  "[cod_bonus] = '"&cod_bonus&"' "
		else
			sql = "SELECT [idBonus_gerado_clientes] " & _
					  ",[Clientes_idClientes] " & _
					  ",[cod_bonus] " & _
					  ",[data_geracao] " & _
					  ",[data_validade] " & _
					  ",[data_resgate] " & _
					  ",[descricao] " & _
					  ",[pontuacao] " & _
					  ",[pontuacao_atingir] " & _
					  ",[numero_solicitacao] " & _
					  ",[moeda] " & _
					  ",[saldo] " & _
					  ",[idproduto] " & _
				  "FROM [marketingoki2].[dbo].[Bonus_Gerado_Clientes] " & _
				  "WHERE [numero_solicitacao] = '"&numero_solicitacao&"' AND " & _
				  "[idproduto] = '"&idproduto&"' AND " & _
				  "[cod_bonus] = '"&cod_bonus&"' "
		end if
'		response.write sql & "<br />"
		call search(sql, arr, intarr)
		if intarr > -1 then
			validaInsertBonusPonto = false
		else
			validaInsertBonusPonto = true
		end if
	end function

	function getSolicitacoesByCliente(id)
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim numero_solicitacao
		dim pontuacao_real
		pontuacao_real = 0
		sql = "select distinct(numero_solicitacao) from bonus_gerado_clientes where clientes_idclientes = " & getCliente(id)
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i=0 then
					numero_solicitacao = arr(0,i)
				else
					if numero_solicitacao <> arr(0,i) then
						numero_solicitacao = arr(0,i)
						pontuacao_real = 0
					end if
				end if
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade) " & _
						"from bonus_gerado_clientes where numero_solicitacao = '"&arr(0,i)&"'"
				call search(sql, arr2, intarr2)
				if intarr > -1 then
					html = html & "<tr>"
					html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
					html = html & "</tr>"
					for j=0 to intarr2
						if j mod 2 = 0 then
							style = "class=""classColorRelPar"""
						else
							style = "class=""classColorRelImpar"""
						end if
						if numero_solicitacao = arr(0,i) then
							pontuacao_real = pontuacao_real + arr2(0,j)
						end if
						html = html & "<tr>"
						html = html & "<td "&style&">"&arr2(0,j)&"</td>"
						html = html & "<td "&style&">"&arr2(1,j)&"</td>"
						html = html & "<td "&style&">"&arr2(2,j)&"/"&arr2(3,j)&"/"&arr2(4,j)&"</td>"
						html = html & "<td "&style&">"&arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j)&"</td>"
						html = html & "</tr>"
					next
					html = html & "<tr>"
					html = html & "<td colspan=""5"" align=""right""><b>Saldo da Pontua��o: </b>"&pontuacao_real&"</td>"
					html = html & "</tr>"
				else
					html = html & "<tr><td colspan=""5"">"&arr(0,i)&"/td></tr>"
				end if
			next
			getSolicitacoesByCliente = html
		else
			getSolicitacoesByCliente = -1
		end if
	end function

	sub atualizaSaldoBonusCliente()
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim numero_solicitacao
		dim pontuacao_real
		dim sql_insert
		dim verificador_bonus_saldo
		pontuacao_real = 0
		sql = "select distinct(numero_solicitacao) from bonus_gerado_clientes"
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i=0 then
					numero_solicitacao = arr(0,i)
				else
					if numero_solicitacao <> arr(0,i) then
						numero_solicitacao = arr(0,i)
						pontuacao_real = 0
					end if
				end if
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo " & _
						"from bonus_gerado_clientes where numero_solicitacao = '"&arr(0,i)&"'"
				call search(sql, arr2, intarr2)
				if intarr2 > -1 then
					for j=0 to intarr2
						verificador_bonus_saldo = arr2(8,j)
						if isempty(verificador_bonus_saldo) or isnull(verificador_bonus_saldo) then
							if numero_solicitacao = arr(0,i) then
								pontuacao_real = pontuacao_real + arr2(0,j)
							end if
						end if
					next
					if pontuacao_real <> 0 then
						sql_insert = "UPDATE [marketingoki2].[dbo].[Bonus_Gerado_Clientes] " & _
									 "SET [saldo] = "&pontuacao_real&" " & _
									 "WHERE [numero_solicitacao] = '"&numero_solicitacao&"'"
	'					response.write "sql update " & sql_insert &	"<br />"
						call exec(sql_insert)
					end if
				end if
			next
		end if
	end sub

	sub atualizaSaldoBonusPonto()
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim numero_solicitacao
		dim pontuacao_real
		dim sql_insert
		dim verificador_bonus_saldo
		pontuacao_real = 0
		sql = "select distinct(numero_solicitacao) from Bonus_Gerado_PontoColeta"
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i=0 then
					numero_solicitacao = arr(0,i)
				else
					if numero_solicitacao <> arr(0,i) then
						numero_solicitacao = arr(0,i)
						pontuacao_real = 0
					end if
				end if
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo " & _
						"from Bonus_Gerado_PontoColeta where numero_solicitacao = '"&arr(0,i)&"'"
				call search(sql, arr2, intarr2)
				if intarr2 > -1 then
					for j=0 to intarr2
						verificador_bonus_saldo = arr2(8,j)
						if isempty(verificador_bonus_saldo) or isnull(verificador_bonus_saldo) then
							if numero_solicitacao = arr(0,i) then
								pontuacao_real = pontuacao_real + arr2(0,j)
							end if
						end if
					next
					if pontuacao_real <> 0 then
						sql_insert = "UPDATE [marketingoki2].[dbo].[Bonus_Gerado_PontoColeta] " & _
									 "SET [saldo] = "&pontuacao_real&" " & _
									 "WHERE [numero_solicitacao] = '"&numero_solicitacao&"'"
	'					response.write "sql update " & sql_insert &	"<br />"
						call exec(sql_insert)
					end if
				end if
			next
		end if
	end sub

	sub atualizaBonusPontoColeta()
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim numero_solicitacao
		dim pontuacao_real
		dim sql_insert
		dim verificador_bonus_saldo
		pontuacao_real = 0
		sql = "select distinct(numero_solicitacao) from bonus_gerado_pontocoleta"
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i=0 then
					numero_solicitacao = arr(0,i)
				else
					if numero_solicitacao <> arr(0,i) then
						numero_solicitacao = arr(0,i)
						pontuacao_real = 0
					end if
				end if
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo " & _
						"from bonus_gerado_pontocoleta where numero_solicitacao = '"&arr(0,i)&"' and saldo <> isnull(saldo, '') "
'				response.write "sql atualiza bonus: " & sql & "<br />"
				call search(sql, arr2, intarr2)
				if intarr2 > -1 then
					for j=0 to intarr2
						verificador_bonus_saldo = arr2(8,j)
						if isempty(verificador_bonus_saldo) or isnull(verificador_bonus_saldo) then
							if numero_solicitacao = arr(0,i) then
								pontuacao_real = pontuacao_real + arr2(0,j)
							end if
						end if
					next
					if pontuacao_real <> 0 then
						sql_insert = "UPDATE [marketingoki2].[dbo].[bonus_gerado_pontocoleta] " & _
									 "SET [saldo] = "&pontuacao_real&" " & _
									 "WHERE [numero_solicitacao] = '"&numero_solicitacao&"'"
	'					response.write "sql update " & sql_insert &	"<br />"
						call exec(sql_insert)
					end if
				end if
			next
		end if
	end sub

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

	sub atualizaKardex()
		dim sql, arr, intarr, i
		dim data
		
		'response.write "entrou"
		'response.end
		
		sql = "select distinct(numero_solicitacao) " & _
		", idproduto " & _
		", pontuacao " & _
		", pontuacao_atingir " & _
		", moeda " & _
		", idcliente " & _
		", data_geracao " & _
		", data_validade " & _
		", data_resgate " & _
		", idbonus " & _
		"from ##valida_gerabonus"
		' numero solicitacao = 0
		' idproduto 		= 1
		' pontuacao 		= 2
		' pontuacao atingir = 3
		' moeda 			= 4
		' idcliente 		= 5
		' data geracao 		= 6
		' data validade 	= 7
		' data resgate 		= 8
		' id bonus 			= 9
		'response.write sql
		'response.end
		call search(sql, arr, intarr)
		
		'response.write "entrou Linha 900" & data
		'response.end  
		
		if intarr > -1 then
			for i=0 to intarr
				if getValidateSolicitacaoBonusGerado(arr(0,i)) = "" or isnull(getValidateSolicitacaoBonusGerado(arr(0,i))) or isempty(getValidateSolicitacaoBonusGerado(arr(0,i))) then
					data = year(now()) & "/" & month(now()) & "/" & day(now())
					sql = "update kardex set data_geracao_bonus = convert(datetime, '"&data&"') where numero_solicitacao_coleta = '"&arr(0,i)&"'"
'					response.write "sql atualiza kardex: " & sql & "<br />"
					call exec(sql)
				end if
			next
		end if
	end sub

	function getDescProduto(idprod)
		dim sql, arr, intarr, i
		sql = "select descricao from produtos where idoki = '"&idprod&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getDescProduto = arr(0,i)
			next
		else
			getDescProduto = ""
		end if
	end function

	function getValidateSolicitacaoBonusGerado(numero)
		dim sql, arr, intarr, i
		sql = "select data_geracao_bonus from kardex where numero_solicitacao_coleta = '"&numero&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getValidateSolicitacaoBonusGerado = arr(0,i)
			next
		end if
	end function

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

	function getProdutoBySolicitacao(num_solicitacao)
		dim sql, arr, intarr
		sql = "select codigo_produto, qtd from kardex " & _
				"inner join Produtos ON IDOki = codigo_produto " & _
				"where numero_solicitacao_coleta = '"&num_solicitacao&"' AND gera_bonus = '1'"
		call search(sql, arr, intarr)
		if intarr > -1 then
'				qtd = trim(arr(1,i))
			getProdutoBySolicitacao = arr
		else
			getProdutoBySolicitacao = ""
		end if
	end function

	sub criaTabelaTemporaria()
		dim sql
		'response.write "TESTE vai criar"
		'response.end
		sql = "create table ##valida_gerabonus ( " & _
					"idcliente int not null, " & _
					"idbonus varchar(50) not null, " & _
					"data_geracao varchar(50) not null, " & _
					"data_validade varchar(50) not null, " & _
					"data_resgate varchar(50) not null, " & _
					"moeda varchar(50) not null, " & _
					"descricao text not null, " & _
					"pontuacao int not null, " & _
					"pontuacao_atingir int not null, " & _
					"numero_solicitacao varchar(13), " & _
					"idproduto varchar(50))"
		'response.write sql
		'response.end
		call exec(sql)
		
		'response.write "Criou"
		'response.end
	end sub

	sub insereTabelaTemporaria(idcliente, idbonus, dataGeracao, dataValidade, dataResgate, moeda, descricao, pontuacao, pontuacaoAtingir, numero_solicitacao, idproduto)
		dim sql
		'response.write "entrou insereTabelaTemporaria"
		'response.end
		if validaInsercaoTabTemp(idcliente, idbonus, numero_solicitacao, idproduto)	then
			'response.write "numero_solicitacao: "&numero_solicitacao&" : produto: "&idproduto&"<br />"
			'response.end
			sql = "insert into ##valida_gerabonus ( " & _
					"idcliente " & _
					",idbonus " & _
					",data_geracao " & _
					",data_validade " & _
					",data_resgate " & _
					",moeda " & _
					",descricao " & _
					",pontuacao " & _
					",pontuacao_atingir " & _
					",numero_solicitacao " & _
					",idproduto) " & _
					"values (" & _
					""&idcliente&" " & _
					",'"&idbonus&"' " & _
					",'"&dataGeracao&"' " & _
					",'"&dataValidade&"' " & _
					",'"&dataResgate&"' " & _
					",'"&moeda&"' " & _
					",'"&descricao&"' " & _
					","&pontuacao&" " & _
					","&pontuacaoAtingir&" " & _
					",'"&numero_solicitacao&"' " & _
					",'"&idproduto&"')"
			'response.write sql & "QUERY<br />"
			'response.end
			call exec(sql)
			'response.write "executou validaInsercaoTabTemp"
			'response.end
		end if
	end sub

	function validaInsercaoTabTemp(idcliente, idbonus, numero_solicitacao, idproduto)
		dim sql, arr, intarr, i
		sql = "select * from ##valida_gerabonus where idcliente = "&idcliente&" " & _
				"and idbonus = '"&idbonus&"' " & _
				"and numero_solicitacao = '"&numero_solicitacao&"' " & _
				"and idproduto = '"&idproduto&"'"
'		response.write sql
		call search(sql, arr, intarr)
		if intarr > -1 then
			'response.write "entrou : Encontrou item na temporaria"
			'response.end
			validaInsercaoTabTemp = false
		else
		    'response.write "Nao encontrou na temporaria"
			'response.end
			validaInsercaoTabTemp = true
		end if
	end function

	sub deletaTabelaTemporaria()
		dim sql
		sql = "drop table ##valida_gerabonus"
		call exec(sql)
	end sub

	function getBonusByCliente(id)
		dim sql, arr, intarr, i
		sql = "select cod_bonus_cli from clientes where idclientes = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getBonusByCliente = arr(0,i)
			next
		else
			getBonusByCliente = -1
		end if
	end function

	function getPontuacaoByProduto(cod_prod, cod_bonus, pontuacao, pontuacaoTarget, qtd)
		dim sql, arr, intarr, i
		sql = "select qtd, pontuacao, pontuacao_target from cadastro_bonus_has_produtos " & _
				"where cad_cod_bonus = '"&trim(cod_bonus)&"' and idoki_prod = '"&trim(cod_prod)&"'"
'		response.write sql & "<br />"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				qtd = arr(0,i)
				pontuacao = arr(1,i)
				pontuacaoTarget = arr(2,i)
				getPontuacaoByProduto = "true"
			next
		else
			qtd = -1
			pontuacao = -1
			pontuacaoTarget = -1
			getPontuacaoByProduto = ""
		end if
	end function

	function getInfoByBonus(idbonus, validade, moeda, descricao)
		dim sql, arr, intarr, i
		sql = "select validade, moeda, descricao from cadastro_bonus where cod_bonus = '"&idbonus&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				validade = arr(0,i)
				moeda = arr(1,i)
				descricao = arr(2,i)
			next
		else
			getValidadeBonus = -1
		end if
	end function

	function getBonusByPontoColeta(id)
		dim sql, arr, intarr, i
		sql = "select bonus_type from pontos_coleta where idpontos_coleta = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getBonusByPontoColeta = arr(0,i)
			next
		else
			getBonusByPontoColeta = -1
		end if
	end function
	'====================================================================================================================

	if request.servervariables("HTTP_METHOD") = "POST" then
		if request.form("action") = "Deletar Selecionados" then
			call deleteSelectedSol()
		else
			logger = ""
			call criaTabelaTemporaria()
			call geraBonusByCliente()
			call deletaTabelaTemporaria()
		end if
	end if
%>
<html>
<head>
<script>
	function checkAll(valor) {
		if (parseInt(document.frmadmkardex.intsol.value) > -1) {
			for (var i=0; i <= parseInt(document.frmadmkardex.intsol.value); i++) {
				var id = "num_"+i;
				if (document.getElementById(id).value == valor) {
					document.getElementById(id).checked = true;
				}
			}
		}
		verifyAnyChecked();
	}

	function verifyAnyChecked() {
		if (parseInt(document.frmadmkardex.intsol.value) > -1) {
			for (var i=0; i <= parseInt(document.frmadmkardex.intsol.value); i++) {
				var id = "num_"+i;
				if (document.getElementById(id).checked) {
					document.getElementById("deleteSelected").style.display = "block";
				}
			}
		}
	}
</script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775" ID="Table1">
		<form action="#" name="frmadmkardex" method="POST" ID="Form1">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellspacing="1" cellpadding="1" width="100%" border=0 ID="Table2">
						<tr>
							<td colspan="2" id="explaintitle" align="center">Administra��o do Kardex</td>
						</tr>
						<tr>
							<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr id="deleteSelected">
							<td>
								<input type="submit" name="action" value="Gerar B�nus" class="btnform" />
							</td>
						</tr>
						<tr>
							<td colspan="4" id="explaintitle" align="left">
								N�mero Solicita��o: <input type="text" name="busca" class="text" value="<%= request.querystring("busca") %>" size="40" />
								<input type="button" name="btnprocurar" value="Procurar" class="btnform" onClick="window.location.href='frmadmkardex.asp?busca=' + document.frmadmkardex.busca.value + ''" />
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;">
									<tr>
										<th align="center">Log de Erros</th>
									</tr>
									<tr>
										<td><%=logger%></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;">
									<tr>
										<th>N� Solicita��o</th>
										<th>C�d. Cliente</th>
										<th>DT. Recebimento</th>
										<th>C�d. Produto</th>
										<th>Desc. Produto</th>
										<th>Quantidade</th>
										<th>DT. Gera��o B�nus</th>
									</tr>
									<%=getListKardex()%>
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
