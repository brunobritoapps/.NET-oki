<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>

<%
	dim saldoTotalBonus
	dim quantidade
	dim idprodutos
	dim valorprodutos

	saldoTotalBonus = getSaldoByCliente(session("IDPonto"))
	if request.servervariables("HTTP_METHOD") = "POST" then
		if request.form("action") = "Resgatar" then
			if getMoeda(session("IDPonto")) = "P" then
				call submitResgatar()
			else
				call submitResgatarMonetario()
			end if
			call geraSolicitacaoResgate()
		end if
	end if

	' Traz pontuação target de um produto
	function getPontTargetByProd(idprod, codbonus)
		dim sql, arr, intarr, i
		dim soma_target
		soma_target = 0
		sql = "select pontuacao_target, qtd from cadastro_bonus_has_produtos where idoki_prod = '"&idprod&"' and cad_cod_bonus = '"&codbonus&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if clng(arr(1,i)) = 1 then
					soma_target = clng(soma_target) + clng(arr(0,i))
				else
					soma_target = clng(soma_target) + (clng(arr(0,i)) / clng(arr(1,i)))
				end if
			next
		end if
		getPontTargetByProd = soma_target
	end function

	sub geraSolicitacaoResgate()
		dim sql
		dim numero_solicitacao
		dim idsolicitacao
		dim arridentity, intidentity, i
		dim arr, intarr, j
		dim idprod, valueprod, somaqtdprodutos, somapontuacaotarget
		somapontuacaotarget = 0
		somaqtdprodutos = 0

		if len(Month(Now())) = 1 then
			numero_solicitacao = "R0"&Month(Now())&Right(Year(Now()), 2)
		else
			numero_solicitacao = "R"&Month(Now())&Right(Year(Now()), 2)
		end if
		numero_solicitacao = numero_solicitacao & getSequencial(False)
		numero_solicitacao = getDigitoControle(numero_solicitacao)

		'Response.Write getMoeda(session("IDPonto")) & "<hr>"
		'Response.End
		
		if getMoeda(session("IDPonto")) = "P" then
			idprod = split(idprodutos, ";")
			valueprod = split(valorprodutos, ";")
			for icont = 0 to ubound(valueprod)
				if valueprod(icont) <> ";" and valueprod(icont) <> "" then
					somaqtdprodutos = somaqtdprodutos + clng(valueprod(icont))
					somapontuacaotarget = clng(somapontuacaotarget) + clng(getPontTargetByProd(idprod(icont), trim(request.form("cod_bonus"))))
					somapontuacaotarget = clng(somapontuacaotarget) * clng(valueprod(icont))
				end if
			next
			if clng(somapontuacaotarget) <= clng(getSaldoByCliente(session("IDPonto"))) then
	'			response.end
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						   "([Status_coleta_idStatus_coleta] " & _
						   ",[numero_solicitacao_coleta] " & _
						   ",[qtd_cartuchos] " & _
						   ",[qtd_cartuchos_recebidos] " & _
						   ",[data_solicitacao] " & _
						   ",[data_aprovacao] " & _
						   ",[data_programada] " & _
						   ",[data_envio_transportadora] " & _
						   ",[data_entrega_pontocoleta] " & _
						   ",[data_recebimento] " & _
						   ",[motivo_status] " & _
						   ",[isMaster]) " & _
					 "VALUES " & _
						   "(1 " & _
						   ",'"&numero_solicitacao&"' " & _
						   ","&somaqtdprodutos&" " & _
						   ",NULL " & _
						   ",convert(datetime, '"&year(now())&"/"&month(now())&"/"&day(now())&"') " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",0)"
				'response.write sql
				'response.end
				call exec(sql)
				sql = "SELECT @@IDENTITY AS id FROM [marketingoki2].[dbo].[Solicitacao_coleta]"
	'			response.write sql
	'			response.end
				call search(sql, arridentity, intidentity)
				if intidentity > -1 then
					for i=0 to intidentity
						idsolicitacao = arridentity(0,i)
					next
				end if
				dim contprod
				for contprod = 0 to ubound(idprod)
					if idprod(contprod) <> ";" and idprod(contprod) <> "" then
						sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_resgate_Ponto] " & _
									   "([cod_bonus] " & _
									   ",[idsolicitacao] " & _
									   ",[documento_baixa] " & _
									   ",[data_baixa] " & _
									   ",[data_solicitacao_resgate] " & _
									   ",[numero_solicitacao_geracao] " & _
									   ",[idproduto] " & _
									   ",[quantidade] " & _
									   ",[idcliente]) " & _
								 "VALUES " & _
									   "('"&trim(request.form("cod_bonus"))&"' " & _
									   ","&idsolicitacao&" " & _
									   ",NULL " & _
									   ",NULL " & _
									   ",convert(datetime, '"&year(now())&"/"&month(now())&"/"&day(now())&"') " & _
									   ",'"&numeroSolicitacaoResgate(idsolicitacao)&"' " & _
									   ",'"&idprod(contprod)&"' " & _
									   ","&valueprod(contprod)&" " & _
									   ","&session("IDPonto")&")"
	'						response.write sql
	'						response.end
							call exec(sql)
					end if
				next
			else
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade"
			end if
		else
			if clng(quantidade) <= clng(getSaldoByCliente(session("IDPonto"))) then
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						   "([Status_coleta_idStatus_coleta] " & _
						   ",[numero_solicitacao_coleta] " & _
						   ",[qtd_cartuchos] " & _
						   ",[qtd_cartuchos_recebidos] " & _
						   ",[data_solicitacao] " & _
						   ",[data_aprovacao] " & _
						   ",[data_programada] " & _
						   ",[data_envio_transportadora] " & _
						   ",[data_entrega_pontocoleta] " & _
						   ",[data_recebimento] " & _
						   ",[motivo_status] " & _
						   ",[isMaster]) " & _
					 "VALUES " & _
						   "(1 " & _
						   ",'"&numero_solicitacao&"' " & _
						   ","&quantidade&" " & _
						   ",NULL " & _
						   ",convert(datetime, '"&year(now())&"/"&month(now())&"/"&day(now())&"') " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",0)"
				'response.write sql
				'response.end
				call exec(sql)
				sql = "SELECT @@IDENTITY AS id FROM [marketingoki2].[dbo].[Solicitacao_coleta]"
	'			response.write sql
		'		response.end
				call search(sql, arridentity, intidentity)
				if intidentity > -1 then
					for i=0 to intidentity
						idsolicitacao = arridentity(0,i)
					next
				end if
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_resgate_Ponto] " & _
							   "([cod_bonus] " & _
							   ",[idsolicitacao] " & _
							   ",[documento_baixa] " & _
							   ",[data_baixa] " & _
							   ",[data_solicitacao_resgate] " & _
							   ",[numero_solicitacao_geracao] " & _
							   ",[idproduto] " & _
							   ",[quantidade] " & _
							   ",[idcliente]) " & _
						 "VALUES " & _
							   "('"&trim(request.form("cod_bonus"))&"' " & _
							   ","&idsolicitacao&" " & _
							   ",NULL " & _
							   ",NULL " & _
							   ",convert(datetime, '"&year(now())&"/"&month(now())&"/"&day(now())&"') " & _
							   ",'"&numeroSolicitacaoResgate(idsolicitacao)&"' " & _
							   ",NULL " & _
							   ","&quantidade&" " & _
							   ","&session("IDPonto")&")"
							'response.write sql
							'response.end
					call exec(sql)
	'			end if
			else
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade"
			end if
		end if
		
		dim arrSolicitacao, intSolicitacao, iSolicitacao
		sql = "select distinct(numero_solicitacao) from bonus_gerado_PontoColeta where Pontos_coleta_idPontos_coleta = " & getCliente(session("IDPonto"))
'		response.write sql
'		response.end
		call search(sql, arrSolicitacao, intSolicitacao)
		
		if intSolicitacao > -1 then
			for iSolicitacao=0 to intSolicitacao
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_Resgate_has_Solicitacao_Composicao] " & _
							   "([numero_resgate] " & _
							   ",[numero_solicitacao]) " & _
						 "VALUES " & _
							   "('"&numero_solicitacao&"' " & _
							   ",'"&arrSolicitacao(0,iSolicitacao)&"')"
'				response.write sql
'				response.end
				call exec(sql)
			next
		end if
		
		call atualizaSaldoBonus()
		'response.write "<script>window.close();</script>"		
		Response.Redirect("frmbonusgeradopontoadm.asp")
	end sub

	sub atualizaSaldoBonus()
		dim sql, arr, intarr, i
		dim arr2, intarr2, j
		dim data_resgate
		dim saldo
		sql = "select distinct(numero_solicitacao) from bonus_gerado_PontoColeta where Pontos_coleta_idPontos_coleta = " & session("IDPonto")
		call search(sql, arr2, intarr2)
		if intarr2 > -1 then
			for j=0 to intarr2
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo, " & _
						"moeda, " & _
						"day(data_resgate), " & _
						"month(data_resgate), " & _
						"year(data_resgate) " & _
						"from bonus_gerado_PontoColeta where numero_solicitacao = '"&arr2(0,j)&"'"
				call search(sql, arr, intarr)
				if intarr > -1 then
					for i=0 to intarr
						data_resgate = arr(10,i)&"/"&arr(11,i)&"/"&arr(12,i)
						saldo = arr(8,i)
						if datediff("d", arr(5,i)&"/"&arr(6,i)&"/"&arr(7,i), now()) < 0 and clng(saldo) <> 0 then 'and len(data_resgate) = 2 
							if j = intarr2 then
								sql = "update bonus_gerado_PontoColeta set saldo = "&request.form("saldo")&", data_resgate = '"&FormatDate(now(),10)&"' where numero_solicitacao = '"&arr2(0,j)&"'"
							else
								sql = "update bonus_gerado_PontoColeta set saldo = 0, data_resgate = '"&FormatDate(now(),10)&"' where numero_solicitacao = '"&arr2(0,j)&"'"
							end if
							call exec(sql)
						end if
					next
				else
				end if
			next
		end if		
	end sub

	function numeroSolicitacaoResgate(id)
		dim sql, arr, intarr, i
		sql = "select numero_solicitacao_coleta from solicitacao_coleta where idsolicitacao_coleta = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				numeroSolicitacaoResgate = arr(0,i)
			next
		else
			numeroSolicitacaoResgate = ""
		end if
	end function

	sub submitResgatar()
		dim sql, arr, intarr, i
		dim valor_pontuacao_atualizado

		sql = "SELECT A.[idoki_prod] " & _
				  ",A.[qtd] " & _
				  ",A.[pontuacao] " & _
				  ",A.[pontuacao_target] " & _
				  ",A.[cad_cod_bonus] " & _
				  ",B.[descricao] " & _
			  "FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
			  "ON A.[idoki_prod] = B.[IDOki] " & _
			  "WHERE A.[cad_cod_bonus] = '" & getCodBonus(session("IDPonto")) & "'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if request.form(arr(0,i)) <> "" and request.form(arr(0,i)) > "0" then
					idprodutos = idprodutos&";"&arr(0,i)
					valorprodutos = valorprodutos&";"&request.form(arr(0,i))
				end if
'				valor_pontuacao_atualizado = atualizarOpcoesResgate(trim(request.form(arr(0,i))), trim(arr(3,i)), valor_pontuacao_atualizado) & "<br />"
'				response.write "IDProduto: " & request.form(arr(0,i)) & "<br />"
'				response.write "Target: " & trim(arr(3,i))& "<br />"
			next
'				response.end
				'response.write "TESTE"
		else
			'//
		end if
	end sub

	sub submitResgatarMonetario()
		quantidade = request.form("quantidade")
	end sub

	function getOpcoesResgatePonto()
		dim sql, arr, intarr, i
		dim html, style
		dim valor_pontuacao_atualizado
		valor_pontuacao_utulizado = 0
		sql = "SELECT A.[idoki_prod] " & _
				  ",A.[qtd] " & _
				  ",A.[pontuacao] " & _
				  ",A.[pontuacao_target] " & _
				  ",A.[cad_cod_bonus] " & _
				  ",B.[descricao] " & _
			  "FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
			  "ON A.[idoki_prod] = B.[IDOki] " & _
			  "WHERE A.[cad_cod_bonus] = '" & getCodBonus(session("IDPonto")) & "'"

'		response.write sql & "<br />"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if

				html = html & "<td "&style&"><input type=""hidden"" name=""intprodutos"" value="""&intarr&""" /></td>"
				valor_pontuacao_atualizado = atualizarOpcoesResgate(trim(request.form(arr(0,i))), trim(arr(3,i)))
				html = html & "<tr>"
				html = html & "<td "&style&">"&arr(0,i)&"</td>"
				html = html & "<td "&style&">"&arr(5,i)&"</td>"
				html = html & "<td "&style&">"&arr(3,i)&"</td>"
				if clng(valor_pontuacao_atualizado) = 0 then
					html = html & "<td "&style&" align=""center""><input type=""text"" name="""&arr(0,i)&""" value="""" class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /></td>"
				else
					html = html & "<td "&style&" align=""center""><input type=""text"" name="""&arr(0,i)&""" value="""&request.form(trim(arr(0,i)))&""" class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /></td>"
				end if
				html = html & "<td "&style&" align=""center""><input type=""text"" name=""utilizados_"&i&""" value="""&valor_pontuacao_atualizado&""" readonly=""true"" class=""text"" size=""15"" style=""color:#999999;"" /></td>"
				html = html & "</tr>"
			next
		else
			html = ""
		end if
		getOpcoesResgatePonto = html
	end function

	function getOpcoesResgateMonetario()
		dim sql, arr, intarr, i
		dim html, style
		dim valor_pontuacao_atualizado

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
			  "FROM [marketingoki2].[dbo].[Bonus_Gerado_PontoColeta] WHERE [Pontos_coleta_idPontos_coleta] = " & session("IDPonto")
'		response.write sql & "<br />"
		call search(sql, arr, intarr)
		if intarr > -1 then
			if i mod 2 = 0 then
				style = "class=""classColorRelPar"""
			else
				style = "class=""classColorRelImpar"""
			end if
			call atualizaOpcoesResgateMonetario(trim(request.form("quantidade")))
			html = html & "<tr>"
			html = html & "<td "&style&" align=""center""><b>Quantidade a Resgatar:</b> <input type=""text"" name=""quantidade"" value="""&request.form("quantidade")&""" class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /></td>"
			html = html & "</tr>"
			i = i + 1
		else
			html = ""
		end if
		getOpcoesResgateMonetario = html
	end function

	function getCliente(id)
		dim sql, arr, intarr, i
		dim retorno

		sql = "SELECT [idPontos_coleta] " & _
			  "FROM [marketingoki2].[dbo].[Pontos_coleta] where [idPontos_coleta] = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				retorno = arr(0,i)
			next
		else
			retorno = -1
		end if
		getCliente = retorno
	end function

	function habilitaResgate()
		if saldoTotalBonus <> getSaldoByCliente(session("IDPonto")) and len(trim(msg)) = 0 then
			habilitaResgate = true
		else
			habilitaResgate = false
		end if
	end function

	function getSaldoByCliente(id)
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim saldo
		dim saldoTotal
		dim sDataResgate
		saldoTotal = 0
		if id > -1 then
			sql = "select distinct(numero_solicitacao) from Bonus_Gerado_PontoColeta where pontos_coleta_idpontos_coleta = " & id
			'response.write sql
			'response.end
			call search(sql, arr, intarr)
			if intarr > -1 then
				for i=0 to intarr
					sql = "select " & _
							"pontuacao, " & _
							"pontuacao_atingir, " & _
							"day(data_geracao) as dia_geracao, " & _
							"month(data_geracao) as mes_geracao, " & _
							"year(data_geracao) as ano_geracao, " & _
							"day(data_validade), " & _
							"month(data_validade), " & _
							"year(data_validade), " & _
							"saldo, " & _
							"moeda, " & _
							"day(data_resgate), " & _
							"month(data_resgate), " & _
							"year(data_resgate) " & _
							"from Bonus_Gerado_PontoColeta where numero_solicitacao = '"&arr(0,i)&"'"
					'response.write sql & "<br />"
					'response.end
					call search(sql, arr2, intarr2)
					if intarr > -1 then
'						response.write datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) & "<br />"
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
'						html = html & "</tr>"
						j=0
						sDataResgate = arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)
						saldo = arr2(8,j)
						if isnull(saldo) then
							saldo=0
						end if
						'						response.write "saldo: "&saldo
						if datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) < 0 and clng(saldo) >= 0 then 'and len(sDataResgate) = 2 
							saldoTotal = saldoTotal + clng(saldo)
						end if
'						response.write saldoTotal
					end if
				next
				getSaldoByCliente = saldoTotal
			else
				getSaldoByCliente = saldoTotal
			end if
		end if
	end function

	function getCodBonus(id)
		dim sql, arr, intarr, i
		sql = "SELECT [bonus_type] FROM [marketingoki2].[dbo].[Pontos_coleta] WHERE [idPontos_coleta] = " & id
'		response.write sql & "<br />"
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getCodBonus = arr(0,i)
			next
		else
			getCodBonus = ""
		end if
	end function

	function getMoeda(id)
		dim sql, arr, intarr, i
		sql = "SELECT [moeda] FROM [marketingoki2].[dbo].[Bonus_Gerado_PontoColeta] WHERE [Pontos_coleta_idPontos_coleta] = " & id
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getMoeda = arr(0,i)
			next
		else
			getMoeda = ""
		end if
	end function

	function atualizarOpcoesResgate(quantidade, pontuacao_target)
		dim tempSaldoBonus
		tempSaldoBonus = saldoTotalBonus
		if quantidade <> "" and quantidade > "0" then
'			response.write "saldoTotalBonus: " & saldoTotalBonus & "<br />"
'			response.write "pontuacao_target: " & pontuacao_target & "<br />"
'			response.write "tempSaldoBonus: " & tempSaldoBonus & "<br />"
'			response.write "valor: "&clng(quantidade) * clng(pontuacao_target)&"<br />"
			tempSaldoBonus = tempSaldoBonus - clng(quantidade) * clng(pontuacao_target)
'			response.write tempSaldoBonus & "<br />"
			if clng(tempSaldoBonus) > 0 then
				saldoTotalBonus = saldoTotalBonus - clng(quantidade) * clng(pontuacao_target)
				atualizarOpcoesResgate = clng(quantidade) * clng(pontuacao_target)
			else
				atualizarOpcoesResgate =  0
'				msg = "O saldo é menor do que o necessário para essa quantidade de produtos"
'				response.write "E menor"
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade de produtos"
			end if
		else
			atualizarOpcoesResgate =  0
			msg = "O saldo é menor do que o necessário para essa quantidade de produtos"
		end if
	end function

	function atualizaOpcoesResgateMonetario(quantidade)
		dim tempSaldoBonus
		tempSaldoBonus = saldoTotalBonus
'		response.write quantidade & "<br />"
'		response.write clng(getSaldoByCliente(session("IDCliente"))) & "<br />"
		if quantidade <> "" and quantidade > "0" then
			tempSaldoBonus = tempSaldoBonus - clng(quantidade)
'			response.write "tempSaldoBonus: " & tempSaldoBonus & "<br />"
'			response.write "tempSaldoBonus: " & saldoTotalBonus & "<br />"
			if clng(tempSaldoBonus) < clng(saldoTotalBonus) and clng(tempSaldoBonus) >= 0 then
				saldoTotalBonus = saldoTotalBonus - clng(quantidade)
				atualizaOpcoesResgateMonetario = clng(quantidade)
				msg = ""
			else
'				response.write "E menor"
'				msg = "O saldo é menor do que o necessário para essa quantidade de produtos"
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade de produtos"
			end if
		else
			atualizaOpcoesResgateMonetario = ""
		end if
	end function
	
	Function FormatDate(data, Forma)
	Dim Dia, Mes, Ano, AnoB, Hora, Minuto, strSemana 
	Dim strPos   
		If Not IsDate(data) Then
			Exit Function	
		End If    

		Dia = "" & Right("00" & Cstr(Day(data)), 2)    
		Mes = "" & Right("00" & Cstr(Month(data)), 2)    
		Ano = "" & Right("0000" & Cstr(Year(data)), 4)
		AnoB = "" & Right("00" & Cstr(Year(data)), 2)
		Hora = "" & Right("00" & Cstr(Hour(data)), 2)    
		Minuto = "" & Right("00" & Cstr(Minute(data)), 2)
		
		Select Case Forma
		Case 1	FormatDate = CStr(Trim(Dia) & "." & Trim(Mes) &"."& Trim(Ano))
		Case 2	FormatDate = CStr(Trim(Ano) & "." & Trim(Mes) &"."& Trim(Dia))
		Case 3	FormatDate = CStr(Trim(Dia) & "." & Trim(Mes) &"."& Trim(Ano) &" - "& Trim(Hora) &":"& Trim(Minuto) &"h")
		Case 4	FormatDate = CStr(Trim(Dia) & "/" & Trim(Mes) &"/"& Trim(Ano))
		Case 5	FormatDate = Trim(Dia) & " " & MonthName(Month(Trim(data))) &" "& Trim(Ano)
		Case 6	FormatDate = CStr(Trim(Dia) & "." & Trim(Mes) &"."& Trim(AnoB))
		Case 7	FormatDate = CStr(Trim(Dia) & "." & Trim(Mes))
		Case 8		
			strSemana = lcase(WeekDayName(WeekDay(data)))
			strPos = len(strSemana)
			if Instr(1, strSemana, "-", 1) > 0 then strPos = (Instr(1, strSemana, "-", 1)-1)
			strSemana = Mid(strSemana,1,strPos)
			FormatDate = Trim(Dia) & "." & MonthName(Month(Trim(data))) &"."& strSemana
		Case 7	FormatDate = CStr(Trim(Dia) & "-" & Trim(Mes) &"-"& Trim(Ano))
		Case 9	FormatDate = CStr(Trim(Dia) & "/" & Trim(Mes) &"/"& Trim(AnoB))
		Case 10	FormatDate = CStr(Trim(Ano) & "-" & Trim(Mes) &"-"& Trim(Dia))
		End Select
	End Function		
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">

<style>
	label {
		font-weight:bold;
	}
</style>

<title><%=TITLE%></title>
<script>
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmviewopcoesresgate.asp" name="frmviewbonusgeradocliente" method="POST">
		<input type="hidden" name="cod_bonus" value="<%=getCodBonus(session("IDPonto"))%>">
		<input type="hidden" name="saldototal" value="<%=getSaldoByCliente(session("IDPonto"))%>">
		<table cellpadding="1" cellspacing="1" width="748" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Visualizar Opções de Resgate<br>ATENÇÃO: Após digitar o valor a ser resgatado, clicar no carrinho e posteriormente no botão Resgatar.</td>
			</tr>
			<tr>
				<td colspan="2" align="right"><a class="linkOperacional" href="frmbonusgeradopontoadm.asp">&laquo Voltar</a></td>
			</tr>
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:615px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<%if getMoeda(session("IDPonto")) = "P" then%>
						<tr>
							<th>Cod Produto</th>
							<th>Descrição do Produto</th>
							<th>Pontuação Target</th>
							<th>Quantidade</th>
							<th>Pontos utilizados</th>
						</tr>
						<%=getOpcoesResgatePonto()%>
						<%else%>
						<tr>
							<th>Quantidade</th>
						</tr>
						<%= getOpcoesResgateMonetario() %>
						<%end if%>
					</table>
					<%=msg%>
					<div id="explaintitle" align="right" style="padding:2px 2px 2px 2px;">
						<%if habilitaResgate() then%>
						<input type="submit" class="btnform" name="action" value="Resgatar" />
						<%end if%>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<b>Saldo do bônus:</b>&nbsp;<input type="text" name="saldo" value="<%= saldoTotalBonus %>" class="text" />
					</div>
					<div align="center"><a href="#" class="linkOperacional"><%= request.querystring("msg") %></a></div>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
