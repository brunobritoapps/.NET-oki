<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%
	dim saldoTotalBonus
	dim quantidade
	dim sSearch
	dim lSaldoAcumulado
	dim arrIdSolicitacoes
	dim steste
	dim valor_pontuacao_atualizado
	dim lItem, lQtdItens
	dim v, i
	dim x, j

	session("ItensPaginacao") = replace(session("ItensPaginacao"), ", , ", "")
	session("ItensPaginacao") = replace(session("ItensPaginacao"), "##", "#")

	if request.servervariables("HTTP_METHOD") = "POST" then
		if request.form("action") = "Resgatar" then
			if getMoeda(session("IDCliente")) = "P" then
				call submitResgatar()
			else
				call submitResgatarMonetario()
			end if
			call geraSolicitacaoResgate()
		end if
		if request.form("action") = "Procurar" then
			sSearch = request("txtSearch")
			call getOpcoesResgatePonto()
		end if
	end if

	function AddItensPaginacao(sItem, lQtd)
		'Response.Write session("ItensPaginacao") & "<hr>"
		'Response.End

		v = split(session("ItensPaginacao"), "#")

		if ubound(v) = -1 then
			session("ItensPaginacao") = sItem & ";" & lQtd & "#" & session("ItensPaginacao")
		else
			if instr(session("ItensPaginacao"), sItem) then
				for i=0 to ubound(v)
					if len(trim(v(i))) > 0 then
						x = split(v(i), ";")
						if x(0) = sItem then
							session("ItensPaginacao") = replace(session("ItensPaginacao"), x(0) & ";" & x(1), sItem & ";" & lQtd)
							lQtdItens = x(1)
						end if
					end if
				next
			else
				session("ItensPaginacao") = sItem & ";" & lQtd & "#" & session("ItensPaginacao")
			end if
		end if

		if lQtd = 0 then

			'Response.Write x(1) & " - " & x(0) & "<hr>"
			session("ItensPaginacao") = replace(session("ItensPaginacao"), sItem & ";0", "")
			session("ItensPaginacao") = replace(session("ItensPaginacao"), "##", "#")
		end if
	end function

	function GetValueItemPaginacao(sItem)
		v = split(session("ItensPaginacao"), "#")

		for i=0 to ubound(v)
			if len(trim(v(i))) > 0 then
				x = split(v(i), ";")
				if x(0) = sItem then
					lQtdItens = x(1)
				end if
			end if
		next
	end function

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
		dim idprod, valueprod, qtdprod
		dim somapontuacaotarget, somaqtdprodutos
		dim icont
		dim lTot
		somapontuacaotarget = 0
		somaqtdprodutos = 0

		if len(Month(Now())) = 1 then
			numero_solicitacao = "R0"&Month(Now())&Right(Year(Now()), 2)
		else
			numero_solicitacao = "R"&Month(Now())&Right(Year(Now()), 2)
		end if
		numero_solicitacao = numero_solicitacao & getSequencial(False)
		numero_solicitacao = getDigitoControle(numero_solicitacao)

		if getMoeda(session("IDCliente")) = "P" then
			idprod = split(session("idprodutos"), ";")
			valueprod = split(session("valorprodutos"), ";")
			for icont = 0 to ubound(valueprod)
				if valueprod(icont) <> ";" and valueprod(icont) <> "" then
					somaqtdprodutos = somaqtdprodutos + clng(valueprod(icont))
					somapontuacaotarget = clng(somapontuacaotarget) + + clng(getPontTargetByProd(idprod(icont), trim(request.form("cod_bonus"))))
					somapontuacaotarget = clng(somapontuacaotarget) * clng(valueprod(icont))
					lTot = somapontuacaotarget + lTot
					somapontuacaotarget = 0
				end if
			next
			somapontuacaotarget = lTot

			'Response.Write session("valorprodutos") & "$<br>"
			'Response.Write session("ItensPaginacao") & "<br>"
			'Response.Write somapontuacaotarget
			'Response.End

			if getSaldoByClienteNew(somapontuacaotarget) = "OK" then
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
						   ",convert(datetime, '"&year(now())&"-"&month(now())&"-"&day(now())&"') " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",NULL " & _
						   ",0)"
		'		response.write sql
		'		response.end
				call exec(sql)
				sql = "SELECT @@IDENTITY AS id FROM [marketingoki2].[dbo].[Solicitacao_coleta]"
		'		response.write sql
		'		response.end
				call search(sql, arridentity, intidentity)

				if intidentity > -1 then
					for i=0 to intidentity
						idsolicitacao = arridentity(0,i)
					next
				end if
				dim contprod
				for contprod = 0 to ubound(idprod)
					if idprod(contprod) <> ";" and idprod(contprod) <> "" then
						sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
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
									   ",convert(datetime, '"&year(now())&"-"&month(now())&"-"&day(now())&"') " & _
									   ",'"&numeroSolicitacaoResgate(idsolicitacao)&"' " & _
									   ",'"&idprod(contprod)&"' " & _
									   ","&valueprod(contprod)&" " & _
									   ","&session("IDCliente")&")"
							'response.write sql
							call exec(sql)
					end if
				next
			else
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade"
			end if
		else
			if getSaldoByClienteNew(somapontuacaotarget) = "OK" then
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
		'		response.write sql
		'		response.end
				call exec(sql)
				sql = "SELECT @@IDENTITY AS id FROM [marketingoki2].[dbo].[Solicitacao_coleta]"
		'		response.write sql
		'		response.end
				call search(sql, arridentity, intidentity)
				if intidentity > -1 then
					for i=0 to intidentity
						idsolicitacao = arridentity(0,i)
					next
				end if
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
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
							   ",'" & getMoedaDesc(session("IDCliente")) & "' " & _
							   ","&quantidade&" " & _
							   ","&session("IDCliente")&")"
	'						response.write sql
					call exec(sql)
	'			end if
			else
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade"
			end if
		end if

		v = split(arrIdSolicitacoes, ";")

		for i=0 to ubound(v)
			sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_Resgate_has_Solicitacao_Composicao] " & _
						   "([numero_resgate] " & _
						   ",[numero_solicitacao]) " & _
					 "VALUES " & _
						   "('"&numero_solicitacao&"' " & _
						   ",'"&v(i)&"')"
			call exec(sql)
		next

		'Aqui q estava juntando os bonus, deixar comentado
		'call atualizaSaldoBonus()

		'response.write "<script>window.close();</script>"
		Response.Redirect("frmviewbonuscliente.asp")
	end sub

	sub atualizaSaldoBonus()
		dim sql, arr, intarr, i
		dim arr2, intarr2, j
		dim data_resgate
		dim saldo
		sql = "select distinct(bonus.numero_solicitacao) " & _
			  "from bonus_gerado_clientes as bonus  " & _
			  "left join clientes as cli " & _
			  "on bonus.clientes_idclientes = cli.idclientes " & _
			  "where bonus.clientes_idclientes in (select idclientes from clientes where cod_cli_consolidador = "&session("IDCliente")&" or idclientes = "&session("IDCliente")&")"

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
						"from bonus_gerado_clientes where numero_solicitacao = '"&arr2(0,j)&"'"

				call search(sql, arr, intarr)
				if intarr > -1 then
					for i=0 to intarr
						'arrumar aqui, ele está zerando os primeiro bonus e somando aos ultimos fazendo um só - Jadilson 03-12-2007
						data_resgate = arr(10,i)&"/"&arr(11,i)&"/"&arr(12,i)
						saldo = arr(8,i)
						if datediff("d", arr(5,i)&"/"&arr(6,i)&"/"&arr(7,i), now()) < 0 and len(data_resgate) = 2 and clng(saldo) <> 0 then
							if j = intarr2 then
								sql = "update bonus_gerado_clientes set saldo = "&request.form("saldo")&" where numero_solicitacao = '"&arr2(0,j)&"'"
							else
								sql = "update bonus_gerado_clientes set saldo = 0 where numero_solicitacao = '"&arr2(0,j)&"'"
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
		dim lTotU

		if len(trim(session("ItensPaginacao"))) > 0 then
			v = split(session("ItensPaginacao"), "#")

			for i=0 to ubound(v)
				if len(trim(v(i))) > 0 then
					x = split(v(i), ";")
					session("idprodutos") = session("idprodutos")&";"&x(0)
					session("valorprodutos") = session("valorprodutos")&";"&x(1)
				end if
			next
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



		if len(trim(sSearch)) > 0 then
			sql = "SELECT A.[idoki_prod] " & _
					  ",A.[qtd] " & _
					  ",A.[pontuacao] " & _
					  ",A.[pontuacao_target] " & _
					  ",A.[cad_cod_bonus] " & _
					  ",B.[descricao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
				  "ON A.[idoki_prod] = B.[IDOki] " & _
				  "WHERE A.[cad_cod_bonus] = '" & getCodBonus(session("IDCliente")) & "' "  & _
				  "AND (idoki_prod LIKE '%"&sSearch&"%' OR descricao LIKE '%"&sSearch&"%')"
		else
			sql = "SELECT A.[idoki_prod] " & _
					  ",A.[qtd] " & _
					  ",A.[pontuacao] " & _
					  ",A.[pontuacao_target] " & _
					  ",A.[cad_cod_bonus] " & _
					  ",B.[descricao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
				  "ON A.[idoki_prod] = B.[IDOki] " & _
				  "WHERE A.[cad_cod_bonus] = '" & getCodBonus(session("IDCliente")) & "'"
		end if
		'response.write sql & "<hr>"
		call search(sql, arr, intarr)

		if intarr > -1 then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = intarr+1 'numero total de registros

			if len(trim(sSearch)) > 0 then
				intPag = 0 'pagina atual da paginacao
			else
				intPag = CInt(Request("pg")) 'pagina atual da paginacao
			end if

			If intPag <= 0 Then intPag = 1
			'if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1

			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)

			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1

			html = html & "<tr><td colspan=9><div id=pag>"
			html = html & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			html = html & "</div></td></tr>"

			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if

				html = html & "<td "&style&"><input type=""hidden"" name=""intprodutos"" value="""&intarr&""" /></td>" & vbcrlf
				html = html & "<tr>"
				html = html & "<td "&style&">"&arr(0,i)&"</td>" & vbcrlf
				html = html & "<td "&style&">"&arr(5,i)&"</td>" & vbcrlf
				html = html & "<td "&style&">"&arr(3,i)&"</td>" & vbcrlf

				'response.write request.form(trim(arr(0,i))) & "<hr>"
				if len(trim(replace(request.form(trim(arr(0,i))),",",""))) then
					AddItensPaginacao arr(0,i), request.form(trim(arr(0,i)))
				end if

				valor_pontuacao_atualizado = atualizarOpcoesResgate(trim(request.form(arr(0,i))), trim(arr(3,i)), arr(0,i))

				if clng(valor_pontuacao_atualizado) = 0 then
					'html = html & "<td "&style&" align=""center""><input type=""text"" id="""&arr(0,i)&""" name="""&arr(0,i)&""" value="""" class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" onKeyPress=""return(soNumeros(this,event));"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_remover.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.getElementById('"&arr(0,i)&"').value='0';document.frmviewbonusgeradocliente.submit()"" /></td>" & vbcrlf
					html = html & "<td "&style&" align=""center""><input type=""text"" id="""&arr(0,i)&""" name="""&arr(0,i)&""" value="""" class=""text"" size=""5"" onKeyPress=""return(soNumeros(this,event));"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_remover.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.getElementById('"&arr(0,i)&"').value='0';document.frmviewbonusgeradocliente.submit()"" /></td>" & vbcrlf
				else
					html = html & "<td "&style&" align=""center""><input type=""text"" name="""&arr(0,i)&""""

					GetValueItemPaginacao arr(0,i)

					if len(trim(lQtdItens)) > 0 then
						html = html & " value=""" & lQtdItens & """"
					else
						html = html & " value=""" & request.form(trim(arr(0,i))) & """"
					end if

					'html = html & " class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" onKeyPress=""return(soNumeros(this,event));"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_remover.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.getElementById('"&arr(0,i)&"').value='0';document.frmviewbonusgeradocliente.submit()"" /></td>" & vbcrlf
					html = html & " class=""text"" size=""5"" onKeyPress=""return(soNumeros(this,event));"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /> <img src=""img/icon_remover.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.getElementById('"&arr(0,i)&"').value='0';document.frmviewbonusgeradocliente.submit()"" /></td>" & vbcrlf
				end if
				html = html & "<td "&style&" align=""center""><input type=""text"" name=""utilizados_"&i&""" value="""&valor_pontuacao_atualizado&""" readonly=""true"" class=""text"" size=""15"" style=""color:#999999;"" /></td>" & vbcrlf
				html = html & "</tr>"
			next
			'html = html & "<tr><td colspan=7>"
			'html = html & Paginacao(iNumPags, intarr, request("pag"), "frmViewOpcoesResgate", Request.ServerVariables("QUERY_STRING"))
			'html = html & "</td></tr>"
		else
			html = ""
		end if
		getOpcoesResgatePonto = html
	end function

	function GetTotalUtilizado()
		dim lTotU
		lTotU = 0
		'session("TotalPontosUtilizados") = 0
		v = split(session("ItensPaginacao"), "#")

		for i=0 to ubound(v)
			if len(trim(v(i))) > 0 then
				x = split(v(i), ";")
					sql = "select pontuacao_target from cadastro_bonus_has_produtos where idoki_prod = '"&x(0)&"' and cad_cod_bonus = '"&getCodBonus(session("IDCliente"))&"'"

					call search(sql, arr, intarr)

					if intarr > -1 then
						lTotU = (arr(0, 0)*x(1)) + lTotU
					end if
			end if
		next
		session("TotalPontosUtilizados") = lTotU
	end function

	function getOpcoesResgateMonetario()
		dim sql, arr, intarr, i
		dim html, style

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
			  "FROM [marketingoki2].[dbo].[Bonus_Gerado_Clientes] where [Clientes_idClientes] = " & session("IDCliente")

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
			html = html & "<td "&style&" align=""center""><b>Quantidade a Resgatar:</b> <input type=""text"" name=""quantidade"" value="""&request.form("quantidade")&""" class=""text"" size=""5"" onBlur=""document.frmviewbonusgeradocliente.submit()"" onKeyPress=""return(soNumeros(this,event));"" /> <img src=""img/icon_carrinho.gif"" align=""absmiddle"" class=""imgexpandeinfo"" onclick=""document.frmviewbonusgeradocliente.submit()"" /></td>"
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

		sql = "SELECT [idClientes] " & _
			  ",[cod_cli_consolidador] " & _
			  "FROM [marketingoki2].[dbo].[Clientes] where idClientes = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if arr(1,i) <> "" and not isnull(arr(1,i)) and not isempty(arr(1,i)) and arr(1,i) > 0 then
					retorno = arr(1,i)
				else
					retorno = arr(0,i)
				end if
			next
		else
			retorno = -1
		end if
		getCliente = retorno
	end function

	function habilitaResgate()
		'Response.Write session("ItensPaginacao") & "<hr>" '<= getSaldoByCliente(session("IDCliente")) 'saldoTotalBonus&"<hr>"&getSaldoByCliente(session("IDCliente"))
		'Response.End
		if saldoTotalBonus > -1 and len(trim(msg)) = 0 and session("ItensPaginacao") <> "#" then
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

		if len(trim(saldoTotalBonus)) then
			saldoTotal = saldoTotalBonus
		else
			saldoTotal = 0
		end if

		sql = "select distinct(bonus.numero_solicitacao) " & _
			  "from bonus_gerado_clientes as bonus  " & _
			  "left join clientes as cli " & _
			  "on bonus.clientes_idclientes = cli.idclientes " & _
			  "where bonus.clientes_idclientes in (select idclientes from clientes where cod_cli_consolidador = "&session("IDCliente")&" or idclientes = "&session("IDCliente")&")"
			'response.write sql & "<hr>"
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
						"from bonus_gerado_clientes where numero_solicitacao = '"&arr(0,i)&"'"
					'response.write sql & "<br />"
					'Response.End
				call search(sql, arr2, intarr2)
				if intarr > -1 then
'						response.write datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) & "<br />"
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
'						html = html & "</tr>"
					j=0
					sDataResgate = arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)
					saldo = arr2(8,j)
					if datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) < 0 then 'and len(sDataResgate) = 2 and clng(saldo) <> 0 then
						saldoTotal = saldoTotal + clng(saldo)
					end if
				end if
			next
			getSaldoByCliente = saldoTotal
		else
			getSaldoByCliente = saldoTotal
		end if
	end function

	function getCodBonus(id)
		dim sql, arr, intarr, i
		sql = "SELECT [cod_bonus_cli] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & id
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
		sql = "SELECT [moeda] FROM [marketingoki2].[dbo].[Bonus_Gerado_Clientes] WHERE [Clientes_idClientes] = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getMoeda = arr(0,i)
			next
		else
			getMoeda = ""
		end if
	end function

	function getMoedaDesc(id)
		dim sql, arr, intarr, i
		sql = "SELECT [moeda] FROM [marketingoki2].[dbo].[Bonus_Gerado_Clientes] WHERE [Clientes_idClientes] = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			if arr(0,i) = "R" then
				getMoedaDesc = "PROD_REAL"
			else
				getMoedaDesc = "PROD_DOLAR"
			end if
		else
			getMoedaDesc = ""
		end if
	end function

	function atualizarOpcoesResgate(quantidade, pontuacao_target, sItem)
		dim tempSaldoBonus
		dim sTxt

		tempSaldoBonus = saldoTotalBonus

		sTxt = session("ItensPaginacao")

		if instr(sTxt, sItem) then
			v = split(sTxt, "#")
			for i=0 to ubound(v)
				if len(trim(v(i))) > 0 then
					x = split(v(i), ";")
					if x(0) = sItem then
						atualizarOpcoesResgate = x(1) * pontuacao_target
						'saldoTotalBonus = saldoTotalBonus - atualizarOpcoesResgate
						lQtdItens = quantidade
					end if
				end if
			next
			'session("TotalPontosUtilizados") = saldoTotalBonus
		else
			'lQtdItens = quantidade

			if quantidade <> "" and quantidade > "0" then
				tempSaldoBonus = tempSaldoBonus - clng(quantidade) * clng(pontuacao_target)
				if clng(tempSaldoBonus) >= 0 then
					'saldoTotalBonus = saldoTotalBonus - clng(quantidade) * clng(pontuacao_target)
					atualizarOpcoesResgate = clng(quantidade) * clng(pontuacao_target)
				else
					atualizarOpcoesResgate =  0
					'msg = "O saldo é menor do que o necessário para essa quantidade de produtos"
					'response.write "E menor"
					response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade de produtos"
				end if
			else
				atualizarOpcoesResgate =  0
				msg = "O saldo é menor do que o necessário para essa quantidade de produtos"
			end if
			if len(trim(sTxt)) = 0 then
				'session("TotalPontosUtilizados") = saldoTotalBonus
			end if
		end if
		'saldoTotalBonus = saldoTotalBonus - session("TotalPontosUtilizados")
	end function

	function atualizaOpcoesResgateMonetario(quantidade)
		dim tempSaldoBonus
		tempSaldoBonus = saldoTotalBonus
		if quantidade <> "" and quantidade > "0" then
			tempSaldoBonus = tempSaldoBonus - clng(quantidade)
			if clng(tempSaldoBonus) < clng(saldoTotalBonus) and clng(tempSaldoBonus) >= 0 then
				'saldoTotalBonus = saldoTotalBonus - clng(quantidade)
				atualizaOpcoesResgateMonetario = clng(quantidade)
				msg = ""
			else
				response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade de produtos"
			end if
		else
			atualizaOpcoesResgateMonetario = ""
		end if
	end function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	Function getSaldoByClienteNew(lTotPontosGastos)
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim saldo
		dim saldoTotal
		dim sDataResgate
		saldoTotal = 0
		sql = "select distinct(bonus.numero_solicitacao) " & _
					"from bonus_gerado_clientes as bonus  " & _
					"left join clientes as cli " & _
					"on bonus.clientes_idclientes = cli.idclientes " & _
					"where bonus.clientes_idclientes in (select idclientes from clientes where cod_cli_consolidador = "&session("IDCliente")&" or idclientes = "&session("IDCliente")&") and saldo > 0 and data_validade >= GETDATE()"

		'Response.Write steste & "<hr>" & clng(lSaldoAcumulado) & "<hr>" & clng(lTotPontosGastos)
		'response.end

		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				select case getPontuacaoBySolicitacao(arr(0,i), lTotPontosGastos)
					case "OK"
						getSaldoByClienteNew = "OK"
						steste = "OK9"
						exit Function
					case "ERRO"
						getSaldoByClienteNew = "ERRO"
						steste = "ERRO9"
						exit Function
					'case "ACUMULOU"
					'	getSaldoByClienteNew = "ACUMULOU"
					'	steste = "ACUMULOU9"
				end select
			next
		else
			getSaldoByClienteNew = "ERRO"
			steste = "ERRO9"
		end if
		Response.Write steste & "<hr>" & clng(lSaldoAcumulado) & "<hr>" & clng(lTotPontosGastos)
		Response.End
	end Function

	function getPontuacaoBySolicitacao(solicitacao, lTotPntGastos)
		dim sql, arr, intarr, i

		'pega o saldo da solicitacao e já valida a data
		sql = "select " & _
					"saldo, data_validade " & _
					"from bonus_gerado_clientes " & _
					"where numero_solicitacao = '"&solicitacao&"' " & _
					"and  saldo > 0 " & _
					"and (SELECT DATEDIFF(DAY, GETDATE(), data_validade)) > -1"

		'Response.Write clng(lSaldoAcumulado)
		'Response.End

		call search(sql, arr, intarr)
		if intarr > -1 then
			arrIdSolicitacoes = arrIdSolicitacoes & solicitacao & ";"
			'for i=0 to intarr
				'Response.Write clng(lTotPontosGastos) &" - "&  clng(arr(0,i))
				'Response.End

				if clng(lSaldoAcumulado) > 0 then
					if (clng(lSaldoAcumulado) + clng(arr(0,i))) > clng(lTotPntGastos) then
						'subtrair do saldo os pontos
						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = " & ((clng(lSaldoAcumulado) + clng(arr(0,i))) - clng(lTotPntGastos)) & _
									", data_resgate = '"&FormatDate(now(),10)&"' WHERE numero_solicitacao = '"&solicitacao&"'"

						call exec(sql)

						getPontuacaoBySolicitacao = "OK"
						steste = "OK"
					elseif clng(arr(0,i)) = clng(lTotPntGastos) then
						'zera o saldo
						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = 0, data_resgate = '"&FormatDate(now(),10)&"' " & _
									"WHERE numero_solicitacao = '"&solicitacao&"'"

						call exec(sql)

						getPontuacaoBySolicitacao = "OK"
						steste = "OK"
					else
						lSaldoAcumulado = (clng(lSaldoAcumulado) + clng(arr(0,i)))

						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = 0, data_resgate = '"&FormatDate(now(),10)&"' " & _
									"WHERE numero_solicitacao = '"&solicitacao&"'"

						call exec(sql)

						if clng(lSaldoAcumulado) >= clng(lTotPntGastos) then
							getPontuacaoBySolicitacao = "OK"
							steste = "OK1"
						else
							getPontuacaoBySolicitacao = "ACUMULOU"
							steste = "ACUMULOU1"
						end if
					end if
				else
					if clng(arr(0,i)) > clng(lTotPntGastos) then
						'subtrair do saldo os pontos
						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = " & (clng(arr(0,i)) - clng(lTotPntGastos)) & _
									", data_resgate = '"&FormatDate(now(),10)&"' WHERE numero_solicitacao = '"&solicitacao&"'"

						call exec(sql)

						getPontuacaoBySolicitacao = "OK"
						steste = "OK"
					elseif clng(arr(0,i)) = clng(lTotPntGastos) then
						'zera o saldo
						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = 0, data_resgate = '"&FormatDate(now(),10)&"' " & _
									"WHERE numero_solicitacao = '"&solicitacao&"'"

						'Response.Write sql
						'Response.End
						call exec(sql)

						getPontuacaoBySolicitacao = "OK"
						steste = "OK"
					else
						lSaldoAcumulado = (clng(lSaldoAcumulado) + clng(arr(0,i)))

						sql = "UPDATE bonus_gerado_clientes " & _
									"SET saldo = 0, data_resgate = '"&FormatDate(now(),10)&"' " & _
									"WHERE numero_solicitacao = '"&solicitacao&"'"

						call exec(sql)

						if clng(lSaldoAcumulado) >= clng(lTotPntGastos) then
							getPontuacaoBySolicitacao = "OK"
							steste = "OK2"
						else
							getPontuacaoBySolicitacao = "ACUMULOU"
							steste = "ACUMULOU2"
						end if
					end if
				end if
			'next
		else
			getPontuacaoBySolicitacao = "ERRO"
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
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<style>
	label {
		font-weight:bold;
	}
</style>
<title><%=TITLE%></title>
<script>
//alert('Escolha os ítens da lista e\n a quantidade e no final clique\n no botão Resgatar no final da página.');
function soNumeros(pFld, e) {
	var sep = 0;
	var key = '';
	var i = j = 0;
	var len = len2 = 0;
	var strCheck = '0123456789';
	var aux = aux2 = '';
	var whichCode = (window.Event) ? e.which : e.keyCode;
	if (whichCode == 13) return true;
	if (whichCode == 0) return true;
	if (whichCode == 8) return true;
	key = String.fromCharCode(whichCode);// Valor para o código da Chave
	if (strCheck.indexOf(key) == -1) return false; // Chave inválida
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmviewopcoesresgate.asp" name="frmviewbonusgeradocliente" method="POST">
		<input type="hidden" name="cod_bonus" value="<%=getCodBonus(session("IDCliente"))%>">
		<input type="hidden" name="saldototal" value="<%=getSaldoByCliente(session("IDCliente"))%>">
		<input type="hidden" name="pg" value="<%If cint(Request("pg")) <= 0 Then Response.Write "1" else Response.Write cint(Request("pg")) end if%>">
		<table cellpadding="1" cellspacing="1" width="748" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Visualizar Opções de Resgate</td>
			</tr>
			<tr>
				<td id="explaintitle" colspan="2" align="center">ATENÇÃO: Após a escolha dos itens da lista e a quantidade, clique no botão Resgatar no final da página.<br>Para adicionar produtos clique no "carrinho" e para zerar clique na "lixeira".</td>
			</tr>
			<tr>
				<td colspan="2" align="right"><a class="linkOperacional" href="frmOperacionalCliente.asp">&laquo Voltar</a></td>
			</tr>
			<tr>
				<td colspan="2">
					Procurar: <INPUT type="text" id=txtSearch name=txtSearch> <input type="submit" class="btnform" name="action" value="Procurar"><br>
					Obs: para listar todos os ítens deixe o campo em branco e clique no botão Procurar
				</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:580px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<%if getMoeda(session("IDCliente")) = "P" then%>
						<tr>
							<th>Cód Produto</th>
							<th>Descrição do Produto</th>
							<th>Pontuação Target</th>
							<th>Quantidade</th>
							<th>Pontos utilizados</th>
						</tr>
						<%
						if len(trim(sSearch)) = 0 then
							if request.form("action") = "Procurar" then
								sSearch = ""
								session("sSearch") = ""
							else
								if len(trim(session("sSearch"))) <> 0 then
									sSearch = session("sSearch")
								else
									sSearch = ""
								end if
							end if
						else
							session("sSearch") = sSearch
						end if
						%>
						<%=getOpcoesResgatePonto()%>
						<%else%>
						<tr>
							<th>Quantidade</th>
						</tr>
						<%= getOpcoesResgateMonetario() %>
						<%end if%>
						<%
						'if len(trim(session("TotalPontosUtilizados"))) = 0 then session("TotalPontosUtilizados") = 0
						call GetTotalUtilizado
						saldoTotalBonus = getSaldoByCliente(session("IDCliente")) - session("TotalPontosUtilizados")

						if clng(saldoTotalBonus) < 0 and  len(trim(request.querystring("msg"))) = 0 then
							response.redirect "frmviewopcoesresgate.asp?msg=O saldo é menor do que o necessário para essa quantidade de produtos"
							'Response.Write "sem saldo"
						end if
						'Response.Write "<hr>"& session("ItensPaginacao") &"<hr>" & session("TotalPontosUtilizados")
						%>
					</table>
					<%=msg%>
					<div id="explaintitle" align="right" style="padding:2px 2px 2px 2px;">
						<%if habilitaResgate() then%>
						<input type="submit" class="btnform" name="action" value="Resgatar" onClick="if(confirm('Tem certeza que selecionou todos os itens desejados?')){return true;}else{return false;}" />
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
