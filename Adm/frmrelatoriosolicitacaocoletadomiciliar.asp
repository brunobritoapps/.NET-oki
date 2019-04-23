<!--#include file="../_config/_config.asp" -->
<%
if request("rm") = "1" then
session("sql") = ""
response.redirect("frmrelatoriosolicitacaocoletadomiciliar.asp")
end if
%>

<%Call open()%>
<%Call GetSessionAdm()%>
<%
	dim tipoSolicitacao
	
    dim pontoColeta
    dim dataEntregaPonto_de
	dim dataEntregaPonto_ate
	dim sqlExportarPonto
	dim sqlExportarCliente

    'COMMENT[1] Brito - Loop -23-10-18 { Variaveis para setar valores retornados na sessao apos consulta }
    dim statusSolicitacao                                                   'A
    dim razaoSocial                                                         'B
	dim transportadora                                                      'C
	dim ufCliente                                                           'D
	dim dataSolicitacao_de                                                  'E
	dim dataSolicitacao_ate                                                 'F
	dim dataAprovacao_de                                                    'G
	dim dataAprovacao_ate                                                   'H
	dim dataRecebimento_de                                                  'I
	dim dataRecebimento_ate                                                 'J
    dim dataProgramada_de                                                   'K
	dim dataProgramada_ate                                                  'L
	
    sub exportarParaArquivo(sql)
		'response.write sql & "</tr>"
		response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/rpttoexcel.aspx?id=" & session("IDCliente") & "&query=" & sql
	end sub

    'INCLUDE[1]Brito-Loop-22-10-18 { Seta VALUES dos campos nas variaveis de sessao apos consulta }
    function getRequest()
        statusSolicitacao = Trim(cint(Request.Form("status")) - 1)          'A
		razaoSocial = Request.Form("razaosocial")                           'B
		transportadora = Trim(Request.Form("transportadora"))               'C
		ufCliente = Trim(Request.Form("uf"))                                'D
		dataSolicitacao_de = Trim(Request.Form("dedatacadastro"))           'E
		dataSolicitacao_ate = Trim(Request.Form("ateatacadastro"))          'F
		dataAprovacao_de = Trim(Request.Form("dedataaprovacao"))            'G
		dataAprovacao_ate = Trim(Request.Form("atedataaprovacao"))          'H
        dataRecebimento_de = Trim(Request.Form("dedatarecebimento"))        'I
        dataRecebimento_ate = Trim(Request.Form("atedatarecebimento"))      'J
		dataProgramada_de = Trim(Request.Form("dedataaprovacao"))           'K
		dataProgramada_ate = Trim(Request.Form("atedataaprovacao"))         'L
	end function

    'OVERRIDE[1]Brito-Loop-22-10-18 { Verifica se foi setado algum valor para uma consulta customizada /se sim retorna TRUE se nao retorna FALSE}
    	function existWhere()
		if  cint(Request.Form("status")) <> 0 or _
			len(Request.Form("razaosocial")) > 0 or _
			cint(Request.Form("transportadora")) <> 0 or _
			Request.Form("uf") <> "0" or _
			len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 or _
			len(Trim(Request.Form("dedataaprovacao"))) > 0 and len(Trim(Request.Form("atedataaprovacao"))) > 0 or _
			len(Trim(Request.Form("dedataprogramada"))) > 0 and len(Trim(Request.Form("atedataprogramada"))) > 0 or _
			len(Trim(Request.Form("dedatarecebimento"))) > 0 and len(Trim(Request.Form("atedatarecebimento"))) > 0 or _
			len(Trim(Request.Form("dedataentrega"))) > 0 and len(Trim(Request.Form("atedataentrega"))) > 0 then
			existWhere = true
		else
			existWhere = false
		end if
	end function

	



    function getSolicitacoesByCliente()
		dim sSql, arr, intarr, i,sNovaQry
		dim html, style

		sNovaQry = "SELECT A.[idSolicitacao_coleta] " & _
				  ",A.[Status_coleta_idStatus_coleta] " & _
				  ",A.[numero_solicitacao_coleta] " & _
				  ",A.[qtd_cartuchos] " & _
				  ",A.[qtd_cartuchos_recebidos] " & _
				  ",A.[data_solicitacao] " & _
				  ",A.[data_aprovacao] " & _
				  ",A.[data_envio_transportadora] " & _
				  ",A.[data_entrega_pontocoleta] " & _
				  ",A.[data_recebimento] " & _
				  ",A.[motivo_status] " & _
				  ",A.[isMaster] " & _
				  ",B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
				  ",B.[typeColect] " & _
				  ",B.[Pontos_coleta_idPontos_coleta] " & _
				  ",B.[Contatos_idContatos] " & _
				  ",B.[Clientes_idClientes] " & _
				  ",B.[cep_coleta] " & _
				  ",B.[logradouro_coleta] " & _
				  ",B.[bairro_coleta] " & _
				  ",B.[numero_endereco_coleta] " & _
				  ",B.[comp_endereco_coleta] " & _
				  ",B.[municipio_coleta] " & _
				  ",B.[estado_coleta] " & _
				  ",B.[ddd_resp_coleta] " & _
				  ",B.[telefone_resp_coleta] " & _
				  ",B.[contato_coleta] " & _
				  ",C.[idClientes] " & _
				  ",C.[Grupos_idGrupos] " & _
				  ",C.[Categorias_idCategorias] " & _
				  ",C.[razao_social] " & _
				  ",C.[nome_fantasia] " & _
				  ",C.[cnpj] " & _
				  ",C.[inscricao_estadual] " & _
				  ",C.[ddd] " & _
				  ",C.[telefone] " & _
				  ",C.[compl_endereco] " & _
				  ",C.[compl_endereco_coleta] " & _
				  ",C.[numero_endereco] " & _
			 	  ",C.[numero_endereco_coleta] " & _
				  ",C.[contato_respcoleta] " & _
				  ",C.[ddd_respcoleta] " & _
				  ",C.[telefone_respcoleta] " & _
				  ",C.[numero_sequencial] " & _
				  ",C.[data_atualizacao_sequencial] " & _
				  ",C.[minCartuchos] " & _
				  ",C.[typeColect] " & _
				  ",C.[status_cliente] " & _
				  ",C.[motivo_status] " & _
				  ",C.[bonus_type] " & _
				  ",C.[Transportadoras_idTransportadoras] " & _
				  ",C.[tipopessoa] " & _
				  ",C.[cod_cli_consolidador] " & _
				  ",C.[cod_bonus_cli] " & _
				  ",A.[data_programada] " & _
			  	  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B " & _
				  "ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Clientes] AS C " & _
				  "ON B.[Clientes_idClientes] = C.[idClientes] " & _
				  "where A.[isMaster] = 0 and left(A.[numero_solicitacao_coleta],1) = 'C' "
        
        sSql = sNovaQry    

		if request.servervariables("HTTP_METHOD") = "POST" then
			call getRequest()
			sSql = sSql & getWhere()

			session("sql") = sSql		
			'if request.form("submit") = "Exportar" then
			'	call exportarParaArquivo(sSql)
			'end if
		else
			if session("sql") <> "" then
				sSql = session("sql")
			end if			
		end if

		'sqlExportarCliente = sSql
        'Response.Write sql & "<hr>"
        
		call search(sSql, arr, intarr)

       if request.form("submit") = "Exportar" then
			call exportarParaArquivo(sSql)
		end if

		if intarr > -1 then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = UBound(arr, 2) + 1 'numero total de registros

			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1

			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)

			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1

			html = html & "<tr><td colspan=10>"
			html = html & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			html = html & "</td></tr>"

			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				html = html & "<tr>"
				html = html & "<td "&style&">"&DateRight(arr(5,i))&"</td>"
				html = html & "<td "&style&">"&DateRight(arr(6,i))&"</td>"
				html = html & "<td "&style&">"&DateRight(arr(54,i))&"</td>"
				html = html & "<td "&style&">"&DateRight(arr(9,i))&"</td>"
				html = html & "<td "&style&">"&arr(2,i)&"</td>"
				html = html & "<td "&style&">"&arr(27,i)&"</td>"
				html = html & "<td "&style&">"&arr(30,i)&"</td>"
				html = html & "<td "&style&">"&arr(23,i)&"</td>"
				html = html & "<td "&style&">"&arr(3,i)&"</td>"
				html = html & "<td "&style&">"&arr(29,i)&"</td>"
				html = html & "<td "&style&">"&getDescCategoria(arr(29,i))&"</td>"
				html = html & "<td "&style&">"&getTransportadoraDesc(arr(50,i))&"</td>"
				html = html & "<td "&style&">"&getStatusDesc(arr(1,i))&"</td>"
				html = html & "</tr>"
			next

		else
			html = html & "<tr><td colspan=8>"
			html = html & "</td></tr>"
			html = html & "<tr>"
			html = html & "<td colspan=""9"" align=""center"">Nenhum registro encontrado</td>"
			html = html & "</tr>"
		end if
		getSolicitacoesByCliente = html
	end function

    'INCLUDE[2]Brito-Loop-22-10-18 { Verifica campos preenchidos conforme validacao existWhere() e concatena query correspondetes aos mesmos}
    function getWhere()
		dim sql
		dim bAnd
		bAnd = false
		if existWhere() then
			bAnd = true
			if cint(Request.Form("status")) <> 0 then
					sql = sql & " and A.[Status_coleta_idStatus_coleta] = " & cint(Request.Form("status"))
			end if
			if len(Request.Form("razaosocial")) > 0 then
					sql = sql & " and C.[razao_social] like '%" & razaoSocial & "%'"
			end if
			if cint(Request.Form("transportadora")) <> 0 then
					sql = sql & " and C.[Transportadoras_idTransportadoras] = " & cint(Request.Form("transportadora"))
			end if
			if Request.Form("uf") <> "0" then
					sql = sql & " and B.[estado_coleta] = '" & Request.Form("uf") & "'"
			end if
			if len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 then
				'if bAnd then
					'sql = sql & " and A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(A.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatacadastro")) & "' and '" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
				'else
				'	'sql = sql & " A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
				'	sql = sql & " (CAST(FLOOR(CAST(A.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatacadastro")) & "' and '" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
				'	bAnd = true
				'end if
			end if
			if len(Trim(Request.Form("dedataaprovacao"))) > 0 and len(Trim(Request.Form("atedataaprovacao"))) > 0 then
			'   if bAnd then
					'sql = sql & " and A.[data_aprovacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedataaprovacao")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(A.[data_aprovacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataaprovacao")) & "' and '" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
			'	else
					'sql = sql & " A.[data_aprovacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedataaprovacao")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
			'		sql = sql & " (CAST(FLOOR(CAST(A.[data_aprovacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataaprovacao")) & "' and '" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
			'		bAnd = true
			'	end if
			end if
			if len(Trim(Request.Form("dedataprogramada"))) > 0 and len(Trim(Request.Form("atedataprogramada"))) > 0 then
			'	if bAnd then
					'sql = sql & " and A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataprogramada")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(A.[data_programada] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataprogramada")) & "' and '" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
			'	else
			'		'sql = sql & " A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataprogramada")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
			'		sql = sql & " (CAST(FLOOR(CAST(A.[data_programada] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataprogramada")) & "' and '" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
			'		bAnd = true
			'	end if
			end if
			if len(Trim(Request.Form("dedatarecebimento"))) > 0 and len(Trim(Request.Form("atedatarecebimento"))) > 0 then
			'	if bAnd then
			'		'sql = sql & " and A.[data_recebimento] between convert(datetime, '" & convertDataSQL(Request.Form("dedatarecebimento")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(A.[data_recebimento] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatarecebimento")) & "' and '" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
			'	else
			'		'sql = sql & " A.[data_recebimento] between convert(datetime, '" & convertDataSQL(Request.Form("dedatarecebimento")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
			'		sql = sql & " (CAST(FLOOR(CAST(A.[data_recebimento] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatarecebimento")) & "' and '" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
			'		bAnd = true
			'	end if
			end if
			'if len(Trim(Request.Form("dedataentrega"))) > 0 and len(Trim(Request.Form("atedataentrega"))) > 0 then
			'	if bAnd then
			'		'sql = sql & " and A.[data_entrega_pontocoleta] between convert(datetime, '" & convertDataSQL(Request.Form("dedataentrega")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataentrega")) & "')"
			'		sql = sql & " and (CAST(FLOOR(CAST(A.[data_entrega_pontocoleta] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataentrega")) & "' and '" & convertDataSQL(Request.Form("atedataentrega")) & "')"
			'	else
			'		'sql = sql & " A.[data_entrega_pontocoleta] between convert(datetime, '" & convertDataSQL(Request.Form("dedataentrega")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataentrega")) & "')"
			'		sql = sql & " (CAST(FLOOR(CAST(A.[data_entrega_pontocoleta] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataentrega")) & "' and '" & convertDataSQL(Request.Form("atedataentrega")) & "')"
			'		bAnd = true
			'	end if
			'end if
		else
			sql = ""
		end if
		getWhere = sql
	end function


	'sub exportarParaArquivo(sqlponto)
	'	dim i, arr, intarr
	'	dim j, arr2, intarr2
	'	dim arquivo
	'	dim fso
	'	dim arquivoPath
	'	dim filenamecsv
	'	dim filename
	'	dim cabecalhoArq
    '    'response.write sqlcliente & "<br />"
	'	'response.write sqlponto
	'	'response.end
	'	set fso = server.createobject("scripting.filesystemobject")
	'	filenamecsv = "exportacao_relatorio_cliente_"&day(now())&"-"&month(now())&"-"&year(now())&"-"&fix(timer())&".csv"
	'	filename = request.servervariables("APPL_PHYSICAL_PATH") & "adm/exportacao/"&filenamecsv
	'	set arquivoPath = fso.createtextfile(filename)
	'	arquivo = ""
	'	call search(sqlponto, arr2, intarr2)

	'	cabecalhoArq = "Data Solicitacao;Data Aprovacao;Data Programada;Data Recebimento;Numero Solicitacao;Cod Cliente;Razao Social;UF Cliente;Qtd. Cartuchos;Cod. Categoria;Desc. Categoria;Transportadora;Status"
	'	arquivoPath.writeLine(cabecalhoArq)

	'	if intarr2 > -1 then
	'		for j=0 to intarr2
	'			arquivo = DateRight(arr2(5,j))&";"&DateRight(arr2(6,j))&";"&DateRight(arr2(54,j))&";"&DateRight(arr2(9,j))&";"&arr2(2,j)&";"&arr2(15,j)&";"&arr2(30,j)&";"&arr2(23,j)&";"&arr2(3,j)&";"&arr2(29,j)&";"&getDescCategoria(arr2(29,j))&";"&getTransportadoraDesc(arr2(50,j))&";"&getStatusDesc(arr2(1,j))
	'			arquivoPath.writeLine(arquivo)
	'		next
	'	end if

	'	'if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
    '    '    response.Write "http://www.sustentabilidadeoki.com.br/adm/exportacao/"&filename
	'		'response.Redirect "http://www.sustentabilidadeoki.com.br/adm/exportacao/"&filenamecsv
	'	'else
	'	'	response.Redirect "http://localhost:81/sgrs/adm/exportacao/"&filenamecsv
	'	'end if

    '    response.Redirect "http://www.sustentabilidadeoki.com.br/adm/exportacao/"&filenamecsv
    '    response.end
	'end sub

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		if isdate(sData) then
			sData = formatdatetime(sData,2)

			dataFormatar = split(sData,"/")
			Dia = replace(dataFormatar(0)," ","")
			Dia = Replace(Dia, "/", "")
			If Len(Dia) = 1 Then
				Dia = "0" & Dia
			End If
			Mes = replace(dataFormatar(1)," ","")
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
			Ano = replace(dataFormatar(2)," ","")
			Ano = Replace(Ano, "/", "")
			Ano = left(Ano, 4)
			DateRight = Mes & "/" & Dia & "/" & Ano
		end if
	End Function

	

	function getStatus()
		dim sql, arr, intarr, i
		dim html
		dim selected

		sql = "SELECT [idStatus_coleta] " & _
				  ",[status_coleta] " & _
			  "FROM [marketingoki2].[dbo].[Status_coleta]"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cint(Request.Form("status")) = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(1,i)&"</option>"
			next
		else
			html = html & "<option value=""0"">---</option>"
		end if
		getStatus = html
	end function

	function getStatusDesc(idStatus)
		dim sql, arr, intarr, i
    
		if len(trim(idStatus)) > 0 then
			sql = "SELECT [Status_coleta] " & _
					  ",[status_coleta] " & _
				  "FROM [marketingoki2].[dbo].[Status_coleta]" & _
				  "WHERE idStatus_coleta = " & idStatus
    
			call search(sql, arr, intarr)
    
			if intarr > -1 then
				getStatusDesc = arr(0,i)
			else
				getStatusDesc = ""
			end if
		end if
	end function

	function getTransportadora()
		dim sql, arr, intarr, i
		dim html, selected

		sql = "SELECT [idTransportadoras] " & _
				  ",[razao_social] " & _
			  "FROM [marketingoki2].[dbo].[Transportadoras]"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cint(Request.Form("transportadora")) = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(1,i)&"</option>"
			next
		else
			html = html & "<option value="""">---</option>"
		end if
		getTransportadora = html
	end function

	function getTransportadoraDesc(idTransportadora)
		dim sql, arr, intarr, i
    
		if len(trim(idTransportadora)) > 0 then
			sql = "SELECT [Razao_social] " & _
					  ",[razao_social] " & _
						"FROM [marketingoki2].[dbo].[Transportadoras] " & _
						"WHERE idTransportadoras = " & idTransportadora
    
			call search(sql, arr, intarr)
    
			if intarr > -1 then
				getTransportadoraDesc = arr(0,i)
			else
				getTransportadoraDesc = ""
			end if
		end if
	end function

	function getUF()
		dim sql, arr, intarr, i
		dim html, selected

		sql = "SELECT distinct([estado]) " & _
			  "FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes]"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if Request.Form("uf") = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(0,i)&"</option>"
			next
		else
			html = html & "<option value="""">---</option>"
		end if
		getUF = html
	end function

	function getPontoColeta()
		dim sql, arr, intarr, i
		dim html, selected
    
		sql = "SELECT [idPontos_coleta] " & _
				  ",[razao_social] " & _
			  "FROM [marketingoki2].[dbo].[Pontos_coleta]"
		call search(sql ,arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cint(Request.Form("pontocoleta")) = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(1,i)&"</option>"
			next
		else
			html = html & "<option value="""">---</option>"
		end if
		getPontoColeta = html
	end function

	function getPontoColetaDesc(idPontoColeta)
		dim sql, arr, intarr, i
    
		if len(trim(idPontoColeta)) > 0 then
			sql = "SELECT [Razao_social] " & _
					  ",[razao_social] " & _
				  "FROM [marketingoki2].[dbo].[Pontos_coleta]"
    
			call search(sql ,arr, intarr)
    
			if intarr > -1 then
				getPontoColetaDesc = arr(0,i)
			else
				getPontoColetaDesc = ""
			end if
		end if
	end function

	function getDescCategoria(id)
		dim sql, arr, intarr, i
		if isempty(id) or isnull(id) then
			getDescCategoria = ""
		else
			sql = "SELECT [descricao] FROM [marketingoki2].[dbo].[Categorias] where [idCategorias] = " & id
			call search(sql, arr, intarr)
			if intarr > -1 then
				for i=0 to intarr
					getDescCategoria = arr(0,i)
				next
			else
				getDescCategoria = ""
			end if
		end if
	end function

	'sub setRequest()
	'	tipoSolicitacao = Trim(Request.Form("tipo"))
	'	statusSolicitacao = Trim(Request.Form("status"))
	'	razaoSocial = Request.Form("razaosocial")
	'	transportadora = Trim(Request.Form("transportadora"))
	'	pontoColeta = Trim(Request.Form("pontocoleta"))
	'	ufCliente = Trim(Request.Form("uf"))
	'	dataSolicitacao_de = Trim(Request.Form("dedatacadastro"))
	'	dataSolicitacao_ate = Trim(Request.Form("atedatacadastro"))
	'	dataAprovacao_de = Trim(Request.Form("dedataaprovacao"))
	'	dataAprovacao_ate = Trim(Request.Form("atedataaprovacao"))
	'	dataProgramada_de = Trim(Request.Form("dedataprogramada"))
	'	dataProgramada_ate = Trim(Request.Form("atedataprogramada"))
	'	dataRecebimento_de = Trim(Request.Form("dedatarecebimento"))
	'	dataRecebimento_ate = Trim(Request.Form("atedatarecebimento"))
	'	dataEntregaPonto_de = Trim(Request.Form("dedataentrega"))
	'	dataEntregaPonto_ate = Trim(Request.Form("atedataentrega"))


    '		 validaDataDeAte(dataSolicitacao_de, dataSolicitacao_ate) & "<br />"
    '========================================================================
    '		Response.Write tipoSolicitacao & "<br />"
    '		Response.Write statusSolicitacao & "<br />"
    '		Response.Write razaoSocial & "<br />"
    '		Response.Write transportadora & "<br />"
    '		Response.Write pontoColeta & "<br />"
    '		Response.Write ufCliente & "<br />"
    '		Response.Write dataSolicitacao_de & "<br />"
    '		Response.Write dataSolicitacao_ate & "<br />"
    '		Response.Write dataAprovacao_de & "<br />"
    '		Response.Write dataAprovacao_ate & "<br />"
    '		Response.Write dataProgramada_de & "<br />"
    '		Response.Write dataProgramada_ate & "<br />"
    '		Response.Write dataRecebimento_de & "<br />"
    '		Response.Write dataRecebimento_ate & "<br />"
    '		Response.Write dataEntregaPonto_de & "<br />"
    '		Response.Write dataEntregaPonto_ate & "<br />"
	'end sub

	function convertDataSQL(data)
		dim splitData
		dim dia, mes, ano
		splitData = split(data,"/")
		if ubound(splitData) > 0 then
			dia = splitData(0)
			mes = splitData(1)
			ano = splitData(2)
			if len(trim(dia)) = 1 then
				dia = "0" & dia
			end if
			if len(trim(mes)) = 1 then
				mes = "0" & mes
			end if
			convertDataSQL = ano & "/" & mes & "/" & dia
		else
			convertDataSQL = ""
		end if
	end function

	function validaDataDeAte(dataDe, dataAte)
		dim validacao
		validacao = datediff("d", dataDe, dataAte)
		if validacao < 0 then
			validaDataDeAte = false
		else
			validaDataDeAte = true
		end if
	end function

	'function existWhereCliente()
	'	if  cint(Request.Form("tipo")) = 2 or cint(Request.Form("tipo")) = 3 or _
	'		cint(Request.Form("status")) <> 0 or _
	'		len(Request.Form("razaosocial")) > 0 or _
	'		cint(Request.Form("transportadora")) <> 0 or _
	'		cint(Request.Form("pontocoleta")) <> 0 or _
	'		Request.Form("uf") <> "0" or _
	'		len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 or _
	'		len(Trim(Request.Form("dedataaprovacao"))) > 0 and len(Trim(Request.Form("atedataaprovacao"))) > 0 or _
	'		len(Trim(Request.Form("dedataprogramada"))) > 0 and len(Trim(Request.Form("atedataprogramada"))) > 0 or _
	'		len(Trim(Request.Form("dedatarecebimento"))) > 0 and len(Trim(Request.Form("atedatarecebimento"))) > 0 or _
	'		len(Trim(Request.Form("dedataentrega"))) > 0 and len(Trim(Request.Form("atedataentrega"))) > 0 then
	'		existWhereCliente = true
	'	else
	'		existWhereCliente = false
	'	end if
	'end function



	

	'function getWhereSQLPonto()
	'	dim sql
	'	dim bAnd
	'	bAnd = false
	'	if existWhere() then
	'		sql = sql & " and "
	'		if cint(Request.Form("status")) <> 0 then
	'			if bAnd then
	'				sql = sql & " and A.[Status_coleta_idStatus_coleta] = " & cint(Request.Form("status"))
	'			else
	'				sql = sql & " A.[Status_coleta_idStatus_coleta] = " & cint(Request.Form("status"))
	'				bAnd = true
	'			end if
	'		end if
	'		if len(Request.Form("razaosocial")) > 0 then
	'			if bAnd then
	'				sql = sql & " and C.[razao_social] like '%" & Request.Form("razaosocial") & "%'"
	'			else
	'				sql = sql & " C.[razao_social] like '%" & Request.Form("razaosocial") & "%'"
	'				bAnd = true
	'			end if
	'		end if
	'		if cint(Request.Form("transportadora")) <> 0 then
	'			if bAnd then
	'				sql = sql & " and C.[idtransp] = " & cint(Request.Form("transportadora"))
	'			else
	'				sql = sql & " C.[idtransp] = " & cint(Request.Form("transportadora"))
	'				bAnd = true
	'			end if
	'		end if
	'		if cint(Request.Form("pontocoleta")) <> 0 then
	'			if bAnd then
	'				sql = sql & " and C.[idPontos_coleta] = " & cint(Request.Form("pontocoleta"))
	'			else
	'				sql = sql & " C.[idPontos_coleta] = " & cint(Request.Form("pontocoleta"))
	'				bAnd = true
	'			end if
	'		end if
	'		if Request.Form("uf") <> "0" then
	'			if bAnd then
	'				sql = sql & " and C.[estado] = '" & Request.Form("uf") & "'"
	'			else
	'				sql = sql & " C.[estado] = '" & Request.Form("uf") & "'"
	'				bAnd = true
	'			end if
	'		end if
	'		if len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 then
	'			if bAnd then
	'				'sql = sql & " and A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
	'				sql = sql & " and (CAST(FLOOR(CAST(A.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatacadastro")) & "' and '" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
	'			else
	'				'sql = sql & " A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
	'				sql = sql & " (CAST(FLOOR(CAST(A.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatacadastro")) & "' and '" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
	'				bAnd = true
	'			end if
	'		end if
	'		if len(Trim(Request.Form("dedataaprovacao"))) > 0 and len(Trim(Request.Form("dedataaprovacao"))) > 0 then
	'			if bAnd then
	'				'sql = sql & " and A.[data_aprovacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedataaprovacao")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
	'				sql = sql & " and (CAST(FLOOR(CAST(A.[data_aprovacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataaprovacao")) & "' and '" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
	'			else
	'				'sql = sql & " A.[data_aprovacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedataaprovacao")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
	'				sql = sql & " (CAST(FLOOR(CAST(A.[data_aprovacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataaprovacao")) & "' and '" & convertDataSQL(Request.Form("atedataaprovacao")) & "')"
	'				bAnd = true
	'			end if
	'		end if
	'		if len(Trim(Request.Form("dedataprogramada"))) > 0 and len(Trim(Request.Form("atedataprogramada"))) > 0 then
	'			if bAnd then
	'				'sql = sql & " and A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataprogramada")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
	'				sql = sql & " and (CAST(FLOOR(CAST(A.[data_programada] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataprogramada")) & "' and '" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
	'			else
	'				'sql = sql & " A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataprogramada")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
	'				sql = sql & " (CAST(FLOOR(CAST(A.[data_programada] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedataprogramada")) & "' and '" & convertDataSQL(Request.Form("atedataprogramada")) & "')"
	'				bAnd = true
	'			end if
	'		end if
	'		if len(Trim(Request.Form("dedatarecebimento"))) > 0 and len(Trim(Request.Form("atedatarecebimento"))) > 0 then
	'			if bAnd then
	'				'sql = sql & " and A.[data_recebimento] between convert(datetime, '" & convertDataSQL(Request.Form("dedatarecebimento")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
	'				sql = sql & " and (CAST(FLOOR(CAST(A.[data_recebimento] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatarecebimento")) & "' and '" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
	'			else
	'				'sql = sql & " A.[data_recebimento] between convert(datetime, '" & convertDataSQL(Request.Form("dedatarecebimento")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
	'				sql = sql & " (CAST(FLOOR(CAST(A.[data_recebimento] AS float)) AS datetime) BETWEEN '" & convertDataSQL(Request.Form("dedatarecebimento")) & "' and '" & convertDataSQL(Request.Form("atedatarecebimento")) & "')"
	'				bAnd = true
	'			end if
	'		end if
	'	else
	'		sql = ""
	'	end if
	'	getWhereSQLPonto = sql
	'end function

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">

<style>
   
</style>
<!-- INCLUDE[3]Brito-Loop-22-10-18{ INCLUSAO DE BIBLIOTECA JQUEY PARA MANIPULACAO DE ELEMENTOS E EVENTOS HTML }-->
<script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-3.3.1.min.js"></script>

<script language='Javascript'>
// **************************************************
// * Autor : Peter M Jordan - uranking@uranking.com *
// * página: www.uranking.com                       *
// **************************************************

// construindo o calendário
function popdate(obj,div,tam,ddd)
{
    if (ddd)
    {
        day = ""
        mmonth = ""
        ano = ""
        c = 1
        char = ""
        for (s=0;s<parseInt(ddd.length);s++)
        {
            char = ddd.substr(s,1)
            if (char == "/")
            {
                c++;
                s++;
                char = ddd.substr(s,1);
            }
            if (c==1) day    += char
            if (c==2) mmonth += char
            if (c==3) ano    += char
        }
        ddd = mmonth + "/" + day + "/" + ano
    }

    if(!ddd) {today = new Date()} else {today = new Date(ddd)}
    date_Form = eval (obj)
    if (date_Form.value == "") { date_Form = new Date()} else {date_Form = new Date(date_Form.value)}

    ano = today.getFullYear();
    mmonth = today.getMonth ();
    day = today.toString ().substr (8,2)

    umonth = new Array ("Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
    days_Feb = (!(ano % 4) ? 29 : 28)
    days = new Array (31, days_Feb, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

    if ((mmonth < 0) || (mmonth > 11))  alert(mmonth)
    if ((mmonth - 1) == -1) {month_prior = 11; year_prior = ano - 1} else {month_prior = mmonth - 1; year_prior = ano}
    if ((mmonth + 1) == 12) {month_next  = 0;  year_next  = ano + 1} else {month_next  = mmonth + 1; year_next  = ano}
    txt  = "<table bgcolor='#efefff' style='border:solid #D90000; border-width:2' cellspacing='0' cellpadding='3' border='0' width='"+tam+"' height='"+tam*1.1 +"'>"
    txt += "<tr bgcolor='#FFFFFF'><td colspan='7' align='center'><table border='0' cellpadding='0' width='100%' bgcolor='#FFFFFF'><tr>"
    txt += "<td width=20% align=center><a href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+((mmonth+1).toString() +"/01/"+(ano-1).toString())+"') class='Cabecalho_Calendario' title='Ano Anterior'><<</a></td>"
    txt += "<td width=20% align=center><a href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+( "01/" + (month_prior+1).toString() + "/" + year_prior.toString())+"') class='Cabecalho_Calendario' title='Mês Anterior'><</a></td>"
    txt += "<td width=20% align=center><a href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+( "01/" + (month_next+1).toString()  + "/" + year_next.toString())+"') class='Cabecalho_Calendario' title='Próximo Mês'>></a></td>"
    txt += "<td width=20% align=center><a href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+((mmonth+1).toString() +"/01/"+(ano+1).toString())+"') class='Cabecalho_Calendario' title='Próximo Ano'>>></a></td>"
    txt += "<td width=20% align=right><a href=javascript:force_close('"+div+"') class='Cabecalho_Calendario' title='Fechar Calendário'><b>X</b></a></td></tr></table></td></tr>"
    txt += "<tr><td colspan='7' align='right' bgcolor='#D90000' class='mes'><a href=javascript:pop_year('"+obj+"','"+div+"','"+tam+"','" + (mmonth+1) + "') class='linkcalendario'>" + ano.toString() + "</a>"
    txt += " <a href=javascript:pop_month('"+obj+"','"+div+"','"+tam+"','" + ano + "') class='linkcalendario'>" + umonth[mmonth] + "</a> <div id='popd' style='position:absolute'></div></td></tr>"
    txt += "<tr bgcolor='#E60000'><td width='14%' class='dia' align=center><b>Dom</b></td><td width='14%' class='dia' align=center><b>Seg</b></td><td width='14%' class='dia' align=center><b>Ter</b></td><td width='14%' class='dia' align=center><b>Qua</b></td><td width='14%' class='dia' align=center><b>Qui</b></td><td width='14%' class='dia' align=center><b>Sex<b></td><td width='14%' class='dia' align=center><b>Sab</b></td></tr>"

    today1 = new Date((mmonth+1).toString() +"/01/"+ano.toString());
    diainicio = today1.getDay () + 1;
    week = d = 1
    start = false;

    for (n=1;n<= 42;n++)
    {
        if (week == 1)  txt += "<tr bgcolor='#efefff' align=center>"
        if (week==diainicio) {start = true}
        if (d > days[mmonth]) {start=false}
        if (start)
        {
            dat = new Date((mmonth+1).toString() + "/" + d + "/" + ano.toString())
            day_dat   = dat.toString().substr(0,10)
            day_today  = date_Form.toString().substr(0,10)
            year_dat  = dat.getFullYear ()
            year_today = date_Form.getFullYear ()
            colorcell = ((day_dat == day_today) && (year_dat == year_today) ? " bgcolor='#FFCC00' " : "" )
            txt += "<td"+colorcell+" align=center><a href=javascript:block('"+  d + "/" + (mmonth+1).toString() + "/" + ano.toString() +"','"+ obj +"','" + div +"') class='data'>"+ d.toString() + "</a></td>"
            d ++
        }
        else
        {
            txt += "<td class='data' align=center> </td>"
        }
        week ++
        if (week == 8)
        {
            week = 1; txt += "</tr>"}
        }
        txt += "</table>"
        div2 = eval (div)
        div2.innerHTML = txt
}

    // função para exibir a janela com os meses
    function pop_month(obj, div, tam, ano)
    {
      txt  = "<table bgcolor='#D90000' border='0' width=80>"
      for (n = 0; n < 12; n++) { txt += "<tr><td align=center><a class='linkcalendario' href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+("01/" + (n+1).toString() + "/" + ano.toString())+"')>" + umonth[n] +"</a></td></tr>" }
      txt += "</table>"
      popd.innerHTML = txt
    }

    // função para exibir a janela com os anos
    function pop_year(obj, div, tam, umonth)
    {
      txt  = "<table bgcolor='#D90000' border='0' width=160>"
      l = 1
      for (n=1991; n<2012; n++)
      {  if (l == 1) txt += "<tr>"
         txt += "<td align=center><a class='linkcalendario' href=javascript:popdate('"+obj+"','"+div+"','"+tam+"','"+(umonth.toString () +"/01/" + n) +"')>" + n + "</a></td>"
         l++
         if (l == 4)
            {txt += "</tr>"; l = 1 }
      }
      txt += "</tr></table>"
      popd.innerHTML = txt
    }

    // função para fechar o calendário
    function force_close(div)
        { div2 = eval (div); div2.innerHTML = ''}

    // função para fechar o calendário e setar a data no campo de data associado
        function block(data, obj, div)
        {
            force_close (div)
            obj2 = eval(obj)
            obj2.value = data
        }

    //INCLUDE[4]Brito-Loop-22-10-18{ Funcao para limpar campos de datas }-->
    function limparCampos() {
        $("input[name='razaosocial']").val("");
        $('#status option').eq(0).prop('selected', true);
        $('#uf option').eq(0).prop('selected', true);
        $('#transportadora option').eq(0).prop('selected', true);

        $("input[name='dedatacadastro']").val("");
        $("input[name='atedatacadastro']").val("");

        $("input[name='dedataaprovacao']").val("");
        $("input[name='atedataaprovacao']").val("");

        $("input[name='dedatarecebimento']").val("");
        $("input[name='atedatarecebimento']").val("");

        $("input[name='dedataaprovacao']").val("");
        $("input[name='atedataaprovacao']").val("");
    }

    //INCLUDE[5]Brito-Loop-22-10-18 { Verifica campos preenchidos conforme validacao existWhere() e concatena query correspondetes aos mesmos}
    //$(document).ready(function () {
    //   $("input[name='razaosocial']").val("");
    //    $('#status option').eq(0).prop('selected', true);
    //     $('#uf option').eq(0).prop('selected', true);
    //    $('#transportadora option').eq(0).prop('selected', true);
    //  });

</script>

<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container" style="width:100%;">
	<table cellspacing="0" cellpadding="0" width="100%">
	<form action="" name="form1" method="POST">
		<tr>
			<td id="conteudo">
				<table cellspacing="3" cellpadding="2" width="100%" border=0>
					<tr>
						<td colspan="3" id="explaintitle" align="center">Relatório de Solicitação de Coleta</td>
					</tr>
					<tr>
						<td colspan="3" align="left"><a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a></td>
					</tr>
					<tr>
						<td colspan="3">
							<table cellpadding="1" cellspacing="1" width="100%">
								<tr>
									<td width="80%">
										<fieldset style="font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif;">
											<legend style="color:#666666;font-weight:bold;">Filtros</legend>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Status:
												<select id="status" name="status" class="select" style="width:200px;">
													<option value="0">[Selecione]</option>
													<%= getStatus() %>
												</select>
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Razão Social:
												<input name="razaosocial" type="text" class="text" value="<%=Request.Form("razaosocial")%>" size="170" />
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Transportadora:
												<select id="transportadora" name="transportadora" class="select" style="width:200px;">
													<option value="0">[Selecione]</option>
													<%= getTransportadora() %>
												</select>
												<!--
												&nbsp;&nbsp;&nbsp;
												Ponto Coleta:
												<select name="pontocoleta" class="select" style="width:300px;">
													<option value="0">[Selecione]</option>
													<%= getPontoColeta() %>
												</select>
												-->
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												UF Cliente:
												<select id="uf" name="uf" class="select" style="width:300px;">
												  <option value="0">[Selecione]</option>
													<%= getUF() %>
											    </select>
												&nbsp;&nbsp;&nbsp;
												Data da Solicitação -
												De: 
                                                    <!--<input type="radio" id="r1" name="myGroup"/>-->
                                                <input name="dedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata1" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro','pop1','150',document.form1.dedatacadastro.value)" /><span id="pop1" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("atedatacadastro"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata2" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro','pop2','150',document.form1.atedatacadastro.value)" /><span id="pop2" style="position:absolute;margin-left:20px;"></span>
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Data da Aprovação -
												De: 
                                                    <!--<input type="radio" id="r2" name="myGroup" />-->
                                                <input name="dedataaprovacao" type="text" class="text" value="<%=Trim(Request.Form("dedataaprovacao"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata3" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedataaprovacao','pop3','150',document.form1.dedataaprovacao.value)" /><span id="pop3" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedataaprovacao" type="text" class="text" value="<%=Trim(Request.Form("atedataaprovacao"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata4" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedataaprovacao','pop4','150',document.form1.atedataaprovacao.value)" /><span id="pop4" style="position:absolute;margin-left:20px;"></span>
												&nbsp;&nbsp;&nbsp;
												Data Programada da Coleta -
												De: 
                                                    <!--<input type="radio" id="r3" name="myGroup" />-->
                                                <input name="dedataprogramada" type="text" class="text" value="<%=Trim(Request.Form("dedataprogramada"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata5" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedataprogramada','pop5','150',document.form1.dedataprogramada.value)" /><span id="pop5" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedataprogramada" type="text" class="text" value="<%=Trim(Request.Form("atedataprogramada"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata6" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedataprogramada','pop6','150',document.form1.atedataprogramada.value)" /><span id="pop6" style="position:absolute;margin-left:20px;"></span>
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Data de Recebimento -
												De:
                                                    <!--<input type="radio" id="r4" name="myGroup" />-->
                                                <input name="dedatarecebimento" type="text" class="text" value="<%=Trim(Request.Form("dedatarecebimento"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata7" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatarecebimento','pop7','150',document.form1.dedatarecebimento.value)" /><span id="pop7" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedatarecebimento" type="text" class="text" value="<%=Trim(Request.Form("atedatarecebimento"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata8" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatarecebimento','pop8','150',document.form1.atedatarecebimento.value)" /><span id="pop8" style="position:absolute;margin-left:20px;"></span>
												<!--
												&nbsp;&nbsp;&nbsp;
												Data de entrega no Ponto de Coleta -
												De: <input name="dedataentrega" type="text" class="text" value="<%=Trim(Request.Form("dedataentrega"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata9" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedataentrega','pop9','150',document.form1.dedataentrega.value)" /><span id="pop9" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedataentrega" type="text" class="text" value="<%=Trim(Request.Form("atedataentrega"))%>" size="13" readonly />
												<input TYPE="button" NAME="btndata0" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedataentrega','pop0','150',document.form1.atedataentrega.value)" /><span id="pop0" style="position:absolute;margin-left:20px;"></span>
												-->
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												<input type="submit" class="btnform" value="Procurar" />
												<input name="submit" type="submit" class="btnform" value="Exportar" />
                                                <input TYPE="button" NAME="btnLimpar3" class="btnform" VALUE="Limpar" Onclick="javascript: limparCampos()" />
											</div>
										</fieldset>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan="3">
							<table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente" style="border:1px solid #000000">
								<tr>
									<th>Data Solicitação</th>
									<th>Data Aprovação</th>
									<th>Data Programada</th>
									<th>Data Recebimento</th>
<!--
'Alteraçào feita por Wea Informática
'Programador: Wellington
'Descrição: Incluída a palavra De coleta, para informar que o número desta solicitaçào é a de coleta referente ao ítem 5 da proposta de correções(INTERFACE ADMINSITRADOR / RELATÓRIOS)
-->
									<th>Número Solicitação de Coleta </th>
									<th>Cod. Cliente</th>
									<th>Razão Social</th>
									<th>UF Cliente</th>
									<th>Qtd. Cartuchos</th>
									<th>Cód. Categoria</th>
									<th>Desc. Categoria</th>
									<th>Transportadora</th>
									<th>Status</th>
								</tr>
								<%= getSolicitacoesByCliente() %>
								
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</form>
	</table>
</div>
</body>
</html>
<%Call close()%>

