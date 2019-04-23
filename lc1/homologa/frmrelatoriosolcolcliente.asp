<!--#include file="_config/_config.asp" -->
<%
if request("rm") = "1" then
    session("sql") = ""
    response.redirect("frmrelatoriosolcolcliente.asp")
end if
%>

<%Call open()%>
<%Call getSessionUser()%>
<%

    dim tipoSolicitacao
	dim statusSolicitacao
	dim razaoSocial
	dim transportadora
	dim pontoColeta
	dim ufCliente
	dim dataSolicitacao_de
	dim dataSolicitacao_ate
	dim dataAprovacao_de
	dim dataAprovacao_ate
	dim dataProgramada_de
	dim dataProgramada_ate
	dim dataRecebimento_de
	dim dataRecebimento_ate
	dim dataEntregaPonto_de
	dim dataEntregaPonto_ate
	
	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano
		
		if sData <> "" then		
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
		else
			DateRight = ""
		end if	
	End Function
	
	sub exportarParaArquivo(sql)
		response.write sql & "</tr>"
		response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/rpttoexcel.aspx?id=" & session("IDCliente") & "&query=" & sql
	end sub
	
	function getDescStatus(id)
		dim sql, arr, intarr, i
		sql = "select status_coleta from status_coleta where idstatus_coleta = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getDescStatus = arr(0,i)
			next
		else
			getDescStatus = ""
		end if
	end function
	
	function getSolicitacoesByCliente()

        dim squery
		dim sql, arr, intarr, i
		dim html, style

        dim scampos

        'incluido peterson 10-5-2014
		dim tipoSolicitacao, statusSolicitacao, dataSolicitacao_de, dataSolicitacao_ate

		scampos = "SELECT A.[idSolicitacao_coleta] " & _
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
				  ",A.[data_programada] "

			scamposExportacao = " " & _ 
            "SELECT   " & _ 
            "convert( varchar(10),A.[data_solicitacao]	,103) as 'Data de Solicitacao do Pedido de Coleta' 	,  " & _ 
            "convert( varchar(10),A.[data_aprovacao],103) as 'Data de Aprovacao da Coleta'	, " & _ 
            "A.[numero_solicitacao_coleta] as 'Numero Solicitacao de Coleta'				, " & _ 
            "' ' as 'Usuario Solicitante',   " & _ 
            "A.[qtd_cartuchos] as 'Qtd. Consumiveis Inserviveis',  " & _ 
            "' ' as 'Qtd de Volumes',   " & _ 
            "' ' as 'Dados para o Faturamento',   " & _ 
            "' ' as 'Dados para a Coleta',   " & _ 
            "IDTRANS.nome_fantasia as 'Transportadora', " & _ 
            "' ' as 'Dt Emissao NF',   " & _ 
            "' ' as 'No Nota Fiscal',   " & _ 
            "convert( varchar(10),A.[data_programada],103) as 'Dt.Programada Coleta no Cliente'	, " & _ 
            "convert( varchar(10),A.[data_recebimento],103) as 'Dt.Coleta no Cliente'		    , " & _ 
            "' ' as 'Dt Chegada no Armazem',   " & _ 

            "A.[qtd_cartuchos_recebidos] as 'Qtd. Consumiveis Inserviveis Recebidos',  " & _ 
            "E.quantidade as 'Qtd Volumes Recebidos',   " & _ 

            "E.Produtos_idProdutos, " & _ 
            "F.descricao, " & _ 
            "D.status_coleta as 'Status' " & _ 
            "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A " & _ 
            "	LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B " & _ 
            "	ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] " & _ 
            "	LEFT JOIN [marketingoki2].[dbo].[Clientes] AS C " & _ 
            "	ON B.[Clientes_idClientes] = C.[idClientes] " & _ 
            "	LEFT JOIN [marketingoki2].[dbo].[status_coleta] AS D " & _ 
            "	ON A.[Status_coleta_idStatus_coleta] = D.idstatus_coleta " & _ 
            "	LEFT JOIN [dbo].[Solicitacoes_coleta_has_Produtos] as E ON E.Solicitacao_coleta_idSolicitacoes_coleta =  A.idSolicitacao_coleta " & _ 
            "	LEFT JOIN [dbo].[Produtos] AS F ON F.IDOki = E.Produtos_idProdutos " & _ 
            "   LEFT JOIN [dbo].[Solicitacao_coleta_has_Transportadoras] AS TRANSP ON TRANSP.Solicitacao_coleta_idSolicitacao_coleta = A.idSolicitacao_coleta " & _ 
            "   LEFT JOIN [dbo].[Transportadoras] AS IDTRANS ON IDTRANS.idTransportadoras = TRANSP.Transportadoras_idTransportadoras "

			if session("cod_cli_consolidador") <> 0 then
				scamposExportacao = scamposExportacao & "where A.[isMaster] = 0 and C.[idClientes] = " & session("IDCliente")
			else
				scamposExportacao = scamposExportacao & "where A.[isMaster] = 0 and C.[idClientes] = " & session("IDCliente")
                '& " or C.[cod_cli_consolidador] = " & session("IDCliente") & ")"
			end if
				  
        sql = scamposExportacao
		sqlExporta = scamposExportacao

        'ver query
		'Response.Write sqlExporta & "<hr>"
		
		if request.ServerVariables("HTTP_METHOD") = "POST" then

			call setRequest()

			if existWhereCliente() then
				sql= sql & getWhereSQLCliente()
				'response.write sql & " <hr>"
			end if
			
            if      request.form("ordem") = "1" Then
                    sql = sql & " " & " order by A.[data_solicitacao],A.[numero_solicitacao_coleta], E.Produtos_idProdutos "
            else    
                if request.form("ordem") = "2" Then
                    sql = sql & " " & " order by A.[data_recebimento],A.[numero_solicitacao_coleta], E.Produtos_idProdutos "
                end if
            end if

			session("sql") = sql

			if request.form("submit") = "Exportar" then
				call exportarParaArquivo(sql)
				'response.write  sqlExporta & "</tr>"
			end if
			
		else
				'response.write "naoque?</tr>"
                if session("sql") <> "" then
                    sql = session("sql")
                else
			        call setRequest()
			        if existWhereCliente() then
    			    	sql= sql & 	getWhereSQLCliente()
	    	    		'response.write sql
            		end if
                end if
                if      request.form("ordem") = "1" Then
                        sql = sql & " " & " order by A.[data_solicitacao],A.[numero_solicitacao_coleta], E.Produtos_idProdutos "
                else 
                    if request.form("ordem") = "2"  then
                        sql = sql & " " & " order by A.[data_recebimento],A.[numero_solicitacao_coleta], E.Produtos_idProdutos "
                    end if
                end if
        end if

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
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
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

				html = html & "<tr>"
				html = html & "<td "&style&">"&arr(00,i)&"</td>"
                html = html & "<td "&style&">"&arr(01,i)&"</td>"
                html = html & "<td "&style&">"&arr(02,i)&"</td>"
                html = html & "<td "&style&">"&arr(03,i)&"</td>"
                html = html & "<td "&style&">"&arr(04,i)&"</td>"
                html = html & "<td "&style&">"&arr(05,i)&"</td>"
                html = html & "<td "&style&">"&arr(06,i)&"</td>"
                html = html & "<td "&style&">"&arr(07,i)&"</td>"
                html = html & "<td "&style&">"&arr(08,i)&"</td>"
                html = html & "<td "&style&">"&arr(09,i)&"</td>"
                html = html & "<td "&style&">"&arr(10,i)&"</td>"
                html = html & "<td "&style&">"&arr(11,i)&"</td>"
                html = html & "<td "&style&">"&arr(12,i)&"</td>"
                html = html & "<td "&style&">"&arr(13,i)&"</td>"
                html = html & "<td "&style&">"&arr(14,i)&"</td>"
                html = html & "<td "&style&">"&arr(15,i)&"</td>"
                html = html & "<td "&style&">"&arr(16,i)&"</td>"
                html = html & "<td "&style&">"&arr(17,i)&"</td>"
                html = html & "<td "&style&">"&arr(18,i)&"</td>"
				'html = html & "<td "&style&">"&getDescStatus(arr(17,i))&"</td>"
				html = html & "</tr>"

			next

			html = html & "<tr><td colspan=9><div id=pag>"
			html = html & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			html = html & "</div></td></tr>"
		else
			html = html & "<tr>"
			html = html & "<td colspan=""9"" align=""center"" class=""classColorRelPar""><b>Nenhum registro encontrado</b></td>"
			html = html & "</tr>"
		end if
	
    	getSolicitacoesByCliente = html

	end function

	function getGrupos() 
		dim sql, arr, intarr, i
		dim html
		html = ""

		'if cint(Session("isMaster")) = 1 then
		if session("cod_cli_consolidador") = 0 then
			sql = "SELECT C.[idClientes] " & _
				  ", C.[nome_fantasia] " & _
			  	  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS B " & _
				  "ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Clientes] AS C " & _
				  "ON B.[Clientes_idClientes] = C.[idClientes] " & _
					"where A.[isMaster] = 0 and C.[cod_cli_consolidador] = " & session("IDCliente") & " OR C.[idClientes] = " & session("IDCliente") & _
					" group by  C.[idClientes], C.[nome_fantasia]"


			'sql = "SELECT    top 1 dbo.Grupos.idGrupos,dbo.Grupos.descricao, dbo.Contatos.Clientes_idClientes, dbo.Contatos.isMaster, dbo.Contatos.usuario, dbo.Contatos.senha " & _
			'"FROM         dbo.Clientes INNER JOIN "  & _
			'                      "dbo.Contatos ON dbo.Clientes.idClientes = dbo.Contatos.Clientes_idClientes INNER JOIN "  & _
			'                      "dbo.Grupos ON dbo.Clientes.Grupos_idGrupos = dbo.Grupos.idGrupos " & _
			'"WHERE     (dbo.Grupos.idGrupos = "&session("GRUPOS_IDGRUPOS")&")"
		'else
			'sql = "SELECT     dbo.Grupos.idGrupos,dbo.Grupos.descricao, dbo.Contatos.Clientes_idClientes, dbo.Contatos.isMaster, dbo.Contatos.usuario, dbo.Contatos.senha " & _
			'"FROM         dbo.Clientes INNER JOIN "  & _
			'                     "dbo.Contatos ON dbo.Clientes.idClientes = dbo.Contatos.Clientes_idClientes INNER JOIN "  & _
			'                      "dbo.Grupos ON dbo.Clientes.Grupos_idGrupos = dbo.Grupos.idGrupos " & _
			'"WHERE     (dbo.Contatos.Clientes_idClientes = "&Session("IDCliente")&")"
				
			'response.write "<hr>" & sql & "<hr>"
			'Response.End
		
			call search(sql, arr, intarr)		  
			if intarr > -1 then
				for i=0 to intarr
					html = html & "<option value="""&arr(0,i)&""">"&arr(1,i)&"</option>"
				next
			end if
		end if
			
		getGrupos = html
	end function

    function getOrdem
        dim select0
        dim select1

        if cint(request.form("ordem")) = 1 then
            select0 = "selected"
        end if

        if cint(request.form("ordem")) = 2 then
            select1 = "selected"
        end if

        html = html & "<option value="""&"1"& """ "&select0&">"&"Dt.Solicitacao"& "</option>"        
        html = html & "<option value="""&"2"& """ "&select1&">"&"Dt.Coleta"& "</option>"
        
        getOrdem = html

    end Function 
	
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
	
	sub setRequest()
		tipoSolicitacao = Trim(Request.Form("tipo"))
		statusSolicitacao = Trim(Request.Form("status"))
		dataSolicitacao_de = Trim(Request.Form("dedatacadastro"))
		dataSolicitacao_ate = Trim(Request.Form("atedatacadastro"))
		
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
	end sub
	
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
	
	function existWhereCliente()
'	response.write "<br>tipo " & cint(Request.Form("tipo"))
'	response.write "<br>status " & cint(Request.Form("status"))
'response.write "<br>deateacastro " &	len(Trim(Request.Form("dedatacadastro"))) 
'response.write "<br>atedatacadastro " & len(Trim(Request.Form("atedatacadastro")))
	
		if  cint(Request.Form("tipo")) > 0 or _
			cint(Request.Form("status")) > 0 or _
			len(Trim(Request.Form("dedatacadastro"))) > 0 or len(Trim(Request.Form("atedatacadastro"))) > 0 then
			existWhereCliente = true
		else
			existWhereCliente = false
		end if	
	end function
	
	function getWhereSQLCliente()
		dim sql
		dim bAnd
		bAnd = false
		
		if existWhereCliente() then
	
			sql = sql & " and "
			if session("ismaster") = 1 then
				if cint(Request.Form("tipo")) <> 0 then	
					sql = sql & " C.[idClientes] = " & cint(Request.Form("tipo"))
					bAnd = true
				end if
			else
			if cint(Request.Form("tipo")) <> 0 then
				sql = sql & " C.[Grupos_idGrupos] = " & cint(Request.Form("tipo"))
				bAnd = true
			end if
			end if

			if cint(Request.Form("status")) <> 0 then

						if bAnd then
							sql = sql & " and A.[Status_coleta_idStatus_coleta] = " & cint(Request.Form("status"))
							bAnd = true
						else
							sql = sql & " A.[Status_coleta_idStatus_coleta] = " & cint(Request.Form("status"))
							bAnd = true
						end if	
			end if
			if len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 then
				if bAnd then
					sql = sql & " and A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & " 00:01') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & " 23:59')"
				else
					sql = sql & " A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & " 00:01') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & " 23:59')"
					bAnd = true
				end if
			end if
		else
			sql = ""
		end if
				
		getWhereSQLCliente = sql
	
	end function
	
%>
<html>
<head>
    <link rel="stylesheet" type="text/css" href="css/geral.css">
    <script language='Javascript'>

        // **************************************************
        // * Autor : Peter M Jordan - uranking@uranking.com *
        // * página: www.uranking.com                       *
        // **************************************************

        // construindo o calendário
        function popdate(obj, div, tam, ddd) {
            if (ddd) {
                day = ""
                mmonth = ""
                ano = ""
                c = 1
                char = ""
                for (s = 0; s < parseInt(ddd.length) ; s++) {
                    char = ddd.substr(s, 1)
                    if (char == "/") {
                        c++;
                        s++;
                        char = ddd.substr(s, 1);
                    }
                    if (c == 1) day += char
                    if (c == 2) mmonth += char
                    if (c == 3) ano += char
                }
                ddd = mmonth + "/" + day + "/" + ano
            }

            if (!ddd) { today = new Date() } else { today = new Date(ddd) }
            date_Form = eval(obj)
            if (date_Form.value == "") { date_Form = new Date() } else { date_Form = new Date(date_Form.value) }

            ano = today.getFullYear();
            mmonth = today.getMonth();
            day = today.toString().substr(8, 2)

            umonth = new Array("Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
            days_Feb = (!(ano % 4) ? 29 : 28)
            days = new Array(31, days_Feb, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

            if ((mmonth < 0) || (mmonth > 11)) alert(mmonth)
            if ((mmonth - 1) == -1) { month_prior = 11; year_prior = ano - 1 } else { month_prior = mmonth - 1; year_prior = ano }
            if ((mmonth + 1) == 12) { month_next = 0; year_next = ano + 1 } else { month_next = mmonth + 1; year_next = ano }
            txt = "<table bgcolor='#efefff' style='border:solid #D90000; border-width:2' cellspacing='0' cellpadding='3' border='0' width='" + tam + "' height='" + tam * 1.1 + "'>"
            txt += "<tr bgcolor='#FFFFFF'><td colspan='7' align='center'><table border='0' cellpadding='0' width='100%' bgcolor='#FFFFFF'><tr>"
            txt += "<td width=20% align=center><a href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + ((mmonth + 1).toString() + "/01/" + (ano - 1).toString()) + "') class='Cabecalho_Calendario' title='Ano Anterior'><<</a></td>"
            txt += "<td width=20% align=center><a href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + ("01/" + (month_prior + 1).toString() + "/" + year_prior.toString()) + "') class='Cabecalho_Calendario' title='Mês Anterior'><</a></td>"
            txt += "<td width=20% align=center><a href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + ("01/" + (month_next + 1).toString() + "/" + year_next.toString()) + "') class='Cabecalho_Calendario' title='Próximo Mês'>></a></td>"
            txt += "<td width=20% align=center><a href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + ((mmonth + 1).toString() + "/01/" + (ano + 1).toString()) + "') class='Cabecalho_Calendario' title='Próximo Ano'>>></a></td>"
            txt += "<td width=20% align=right><a href=javascript:force_close('" + div + "') class='Cabecalho_Calendario' title='Fechar Calendário'><b>X</b></a></td></tr></table></td></tr>"
            txt += "<tr><td colspan='7' align='right' bgcolor='#D90000' class='mes'><a href=javascript:pop_year('" + obj + "','" + div + "','" + tam + "','" + (mmonth + 1) + "') class='linkcalendario'>" + ano.toString() + "</a>"
            txt += " <a href=javascript:pop_month('" + obj + "','" + div + "','" + tam + "','" + ano + "') class='linkcalendario'>" + umonth[mmonth] + "</a> <div id='popd' style='position:absolute'></div></td></tr>"
            txt += "<tr bgcolor='#E60000'><td width='14%' class='dia' align=center><b>Dom</b></td><td width='14%' class='dia' align=center><b>Seg</b></td><td width='14%' class='dia' align=center><b>Ter</b></td><td width='14%' class='dia' align=center><b>Qua</b></td><td width='14%' class='dia' align=center><b>Qui</b></td><td width='14%' class='dia' align=center><b>Sex<b></td><td width='14%' class='dia' align=center><b>Sab</b></td></tr>"

            today1 = new Date((mmonth + 1).toString() + "/01/" + ano.toString());
            diainicio = today1.getDay() + 1;
            week = d = 1
            start = false;

            for (n = 1; n <= 42; n++) {
                if (week == 1) txt += "<tr bgcolor='#efefff' align=center>"
                if (week == diainicio) { start = true }
                if (d > days[mmonth]) { start = false }
                if (start) {
                    dat = new Date((mmonth + 1).toString() + "/" + d + "/" + ano.toString())
                    day_dat = dat.toString().substr(0, 10)
                    day_today = date_Form.toString().substr(0, 10)
                    year_dat = dat.getFullYear()
                    year_today = date_Form.getFullYear()
                    colorcell = ((day_dat == day_today) && (year_dat == year_today) ? " bgcolor='#FFCC00' " : "")
                    txt += "<td" + colorcell + " align=center><a href=javascript:block('" + d + "/" + (mmonth + 1).toString() + "/" + ano.toString() + "','" + obj + "','" + div + "') class='data'>" + d.toString() + "</a></td>"
                    d++
                }
                else {
                    txt += "<td class='data' align=center> </td>"
                }
                week++
                if (week == 8) {
                    week = 1; txt += "</tr>"
                }
            }
            txt += "</table>"
            div2 = eval(div)
            div2.innerHTML = txt
        }

        // funcao para exibir a janela com os meses
        function pop_month(obj, div, tam, ano) {
            txt = "<table bgcolor='#D90000' border='0' width=80>"
            for (n = 0; n < 12; n++) { txt += "<tr><td align=center><a class='linkcalendario' href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + ("01/" + (n + 1).toString() + "/" + ano.toString()) + "')>" + umonth[n] + "</a></td></tr>" }
            txt += "</table>"
            popd.innerHTML = txt
        }

        // funcao para exibir a janela com os anos
        function pop_year(obj, div, tam, umonth) {
            txt = "<table bgcolor='#D90000' border='0' width=160>"
            l = 1
            for (n = 1991; n < 2012; n++) {
                if (l == 1) txt += "<tr>"
                txt += "<td align=center><a class='linkcalendario' href=javascript:popdate('" + obj + "','" + div + "','" + tam + "','" + (umonth.toString() + "/01/" + n) + "')>" + n + "</a></td>"
                l++
                if (l == 4)
                { txt += "</tr>"; l = 1 }
            }
            txt += "</tr></table>"
            popd.innerHTML = txt
        }

        // funcao para fechar o calendário
        function force_close(div)
        { div2 = eval(div); div2.innerHTML = '' }

        // funcao para fechar o calendário e setar a data no campo de data associado
        function block(data, obj, div) {
            force_close(div)
            obj2 = eval(obj)
            obj2.value = data
        }
    </script>
    <title><%=TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style type="text/css">
        .auto-style1 {
            width: 211px;
        }

        .auto-style2 {
            width: 247px;
        }

        .auto-style3 {
            width: 200px;
            text-align: right;
        }
    </style>
</head>

<body>
    <div id="container" style="width: 100%;">
        <table cellspacing="0" cellpadding="0" width="100%">
            <form action="" name="form1" method="POST">
                <tr>
                    <td id="conteudo">
                        <table cellspacing="3" cellpadding="2" width="100%" border="0">
                            <tr>
                                <td id="explaintitle" align="center">Relatório de Solicitacao de Coleta</td>
                            </tr>
                            <tr>
                                <td>
                                    <table cellpadding="1" cellspacing="1" width="100%">
                                        <tr>
                                            <td class="auto-style3">&nbsp;</td>
                                            <td class="auto-style1">&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td class="auto-style2">&nbsp;</td>
                                            <td colspan="2">&nbsp;</td>
                                            <td style="text-align: right;">

                                                <a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a></td>

                                        </tr>
                                        <tr>
                                            <td class="auto-style3">Grupo:</td>
                                            <td class="auto-style1">
                                                <select name="tipo" class="select" style="width: 200px;">
                                                    <option value="0">[Selecione]</option>
                                                    <%=getGrupos()%>
                                                </select>
                                            </td>
                                            <td style="text-align: right">Status:</td>
                                            <td class="auto-style2">
                                                <select name="status" class="select" style="width: 200px;">
                                                    <option value="0">[Selecione]</option>
                                                    <%=getStatus()%>
                                                </select>
                                            </td>
                                            <td colspan="2"></td>
                                            <td></td>
                                        </tr>
                                        <tr>
                                            <td class="auto-style3">Data da Solicitacao De:</td>
                                            <td class="auto-style1">
                                                <input name="dedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro"))%>" size="13" readonly />
                                                <input type="button" name="btndata1" class="btnform" value="..." onclick="javascript: popdate('document.form1.dedatacadastro', 'pop1', '150', document.form1.dedatacadastro.value)" /><span id="pop1" style="position: absolute; margin-left: 20px;"></span>
                                            </td>
                                            <td style="text-align: right">Ate:</td>
                                            <td class="auto-style2">

                                                <input name="atedatacadastro" type="text" class="text" value='<%=Trim(Request.Form("atedatacadastro"))%>' size="13" readonly />
                                                <input class="btnform" name="btndata2" onclick="javascript: popdate('document.form1.atedatacadastro', 'pop2', '150', document.form1.atedatacadastro.value)" type="button" value="..." /><span id="pop2" style="position: absolute; margin-left: 20px;"></span>
                                            </td>
                                            <td colspan="2">&nbsp;</td>
                                            <td>&nbsp;</td>

                                        </tr>
                                        <tr>
                                            <td class="auto-style3">Usuário Solicitante:</td>
                                            <td class="auto-style1">
                                                <input type="text" size="20" maxlength="50" style="text-transform: uppercase;" class="textreadonly" /></td>
                                            <td>&nbsp;</td>
                                            <td class="auto-style2">&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>

                                        </tr>
                                        <tr>
                                            <td class="auto-style3">Ordenar Por:</td>
                                            <td class="auto-style1">
                                                <select name="Ordem" class="select" style="width: 200px;">
                                                    <option value="0">[Selecione]</option>
                                                    <%=getOrdem()%>
                                                </select></td>
                                            <td>&nbsp;</td>
                                            <td class="auto-style2">&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>

                                        </tr>
                                        <tr>
                                            <td class="auto-style3">&nbsp;</td>
                                            <td class="auto-style1">&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td class="auto-style2">&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td style="text-align: right;">&nbsp;<input name="Procurar" class="btnform" type="submit" value="Procurar" />
                                                <input class="btnform" name="submit" type="submit" value="Exportar" /></td>

                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente" style="border: 1px solid #000000">
                                        <tr>
                                            <th>Data de Solicitacao do Pedido de Coleta</th>
                                            <th>Data de Aprovacao da Coleta</th>
                                            <th>Número Solicitacao de Coleta</th>
                                            <th>Usuário Solicitante</th>
                                            <th>Qtd. Consumiveis Inserviveis	</th>
                                            <th>Qtd. De Volumes	</th>
                                            <th>Dados para Faturamento	</th>
                                            <th>Dados para a Coleta	</th>
                                            <th>Transportadora	</th>
                                            <th>Data da Emissao da Nota Fiscal de Coleta	</th>
                                            <th>N° Nota Fiscal da Coleta	</th>
                                            <th>Data Programada para a Coleta no Cliente	</th>
                                            <th>Data da Coleta no Cliente	</th>
                                            <th>Data de Chegada no Armazem	</th>
                                            <th>Qtd. Consumiveis Inserviveis Recebidos	</th>
                                            <th>Qtd. De Volumes Recebidos	</th>
                                            <th>Códigos Recebidos	</th>
                                            <th>Descricao	</th>
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