<!--#include file="../_config/_config.asp" -->
<%
if request("rm") = "1" then
session("sql") = ""
response.redirect("frmrelatoriobonuscliente.asp")
end if
%>
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	dim statusBonus
	dim razaoSocial
	dim dataGeracao_de
	dim dataGeracao_ate
	dim dataExpiracao_de
	dim dataExpiracao_ate
	dim dataResgate_de
	dim dataResgate_ate
	dim descBonus
	dim ufPonto
    
    sub exportarParaArquivo(sql)
		response.write sql & "</tr>"
		response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/rpttoexcel.aspx?id=" & session("IDCliente") & "&query=" & sql
	end sub
    	
	'sub exportarParaArquivo(sql)
	'
	''response.write sql
	''response.end
	'dim i, arr, intarr
	'dim arquivo
	'dim fso
	'dim arquivoPath
	'dim filenamecsv
	'dim filename
	'dim cabecalhoArq
	'
	'set fso = server.createobject("scripting.filesystemobject")
	'filenamecsv = "exportacao_relatorio_cliente_"&day(now())&"-"&month(now())&"-"&year(now())&"-"&fix(timer())&".csv"
	'filename = request.servervariables("APPL_PHYSICAL_PATH") & "adm/exportacao/"&filenamecsv
	'set arquivoPath = fso.createtextfile(filename)
	'arquivo = ""
	'call search(sql, arr, intarr)
	'if intarr > -1 then 
	'	cabecalhoArq = "Código Cliente;Razão Social;Pontuação;Cód. Bônus;Moeda do Bônus;Data Geração;Data Expiração"
	'	arquivoPath.writeLine(cabecalhoArq)
	'	for i=0 to intarr
	'		arquivo = arr(0,i)&";"&arr(8,i)&";"&arr(5,i)&";"&arr(1,i)&";"&arr(7,i)&";"&DateRight(arr(2,i))&";"&DateRight(arr(3,i))
	'		arquivoPath.writeLine(arquivo)
	'	next
	'end if 
	'response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/rpttoexcel.aspx?id=" & session("IDCliente") & "&query=" & sql
	''if left(request.servervariables("LOCAL_ADDR"), 2) = "10" then
	''	response.Redirect "http://www.sustentabilidadeoki.com.br/adm/exportacao/"&filenamecsv
	''else
	''	response.Redirect "exportacao"&filenamecsv
	''end if
	'end sub
	
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

	function getBonusCliente()
		dim sql, arr, intarr, i, sSql
		dim html, style		

        sSql = " " & _ 

			"SELECT distinct  " & _ 
			"SOLCOL.numero_solicitacao_coleta as 'Numero Solicitação' " & _ 
			",'' as 'Usuário Solicitante' " & _ 
			",convert(varchar,B.descricao) " & _ 
			",CLI.[razao_social] as 'Razao Social/Nome' " & _ 
			",CLI.[nome_fantasia] as 'Nome Fantasia' " & _ 
			",CLI.[cnpj] as 'CNPJ/CPF' " & _ 
			",CLI.[inscricao_estadual] as 'IE' " & _ 
			",E.cep as 'CEP' " & _ 
			",E.logradouro as 'Logradouro' " & _ 
			",convert(varchar,CLI.[compl_endereco]) as 'Complemento Logradouro' " & _ 
			",CLI.[numero_endereco] as 'No.' " & _ 
			",E.bairro as 'Bairro' " & _ 
			",E.municipio as 'Municpio' " & _ 
			",E.estado as 'Estado' " & _ 
			",CLI.[ddd] " & _ 
			",CLI.[telefone] " & _ 
			",SOLCOL.[qtd_cartuchos] as 'Qtd Consumíveis Inservíveis' " & _ 
			",SOLhPROD.quantidade as 'Qtd Volumes Recebidos' " & _ 
			",' ' as 'Dt.NF' " & _ 
			",' ' as 'No.NF' " & _ 
			",convert( nvarchar(10), SOLCOL.[data_recebimento],103) as 'Dt.Coleta' " & _ 
			",convert( nvarchar(10), SOLCOL.[data_recebimento],103) as 'Dt.Chegada Armazém' " & _ 
			",SOLhPROD.quantidade as 'Qtd. Inservíveis Recebidos / Conferidos' " & _ 
			",SOLhPROD.quantidade as 'Qtd. De Volumes Recebidos / Conferidos' " & _ 
			",SOLhPROD.Produtos_idProdutos as 'Código Recebido' " & _ 
			",PROD.Grupo_produtos_idGrupo_produtos as 'Grupo de Produtos' " & _ 
			",PROD.[descricao] as 'Descricao do Produto' " & _ 
			",BONUSGER.pontuacao as 'Pontuacao Unitaria' " & _ 
			",BONUSGER.pontuacao * SOLhPROD.quantidade as 'Pontuacao Total' " & _ 
			",convert( nvarchar(10), BONUSGER.data_geracao, 103) as 'Dt.Geracao do Bônus' " & _ 
			",'' as Status " & _ 
			"FROM [Bonus_Gerado_Clientes] as BONUSGER " & _ 
			"	inner join [Solicitacao_coleta] as SOLCOL on BONUSGER.numero_solicitacao = SOLCOL.numero_solicitacao_coleta " & _ 
			"	inner join [Solicitacoes_coleta_has_Produtos] as SOLhPROD ON SOLCOL.idSolicitacao_coleta = SOLhPROD.Solicitacao_coleta_idSolicitacoes_coleta " & _ 
			"	inner join [Solicitacao_coleta_has_Clientes] as SOLCOLCLI ON SOLCOLCLI.Solicitacao_coleta_idSolicitacao_coleta = SOLhPROD.Solicitacao_coleta_idSolicitacoes_coleta " & _ 
			"	inner join [Clientes] as CLI on CLI.idClientes = SOLCOLCLI.Clientes_idClientes  " & _ 
			"	left outer join [marketingoki2].[dbo].[Produtos] as PROD on PROD.IDOki = SOLhPROD.Produtos_idProdutos " & _ 
			"	LEFT JOIN [marketingoki2].[dbo].[Categorias] AS B ON CLI.[Categorias_idCategorias] = B.[idCategorias] " & _ 
			"	LEFT JOIN [marketingoki2].[dbo].[Grupos] AS C ON CLI.[Grupos_idGrupos] = C.[idGrupos] " & _ 
			"	LEFT JOIN [marketingoki2].[dbo].[cadastro_bonus] as D ON  D.cod_bonus = CLI.cod_bonus_cli  " & _ 
			"	LEFT JOIN lc_cep_consulta_has_Clientes AS E on CLI.idClientes = E.Clientes_idClientes and E.isEnderecoComum = 1 " & _ 
			"	where BONUSGER.Clientes_idClientes = CLI.idClientes and BONUSGER.numero_solicitacao = SOLCOL.numero_solicitacao_coleta and SOLhPROD.Produtos_idProdutos = BONUSGER.idproduto " 

                
        sql = sSql
	
		if request.servervariables("HTTP_METHOD") = "POST" then
			call setRequest()
			sql = sql & getWhereSQL()
			'sql = sql & " GROUP BY A.numero_solicitacao, A.Clientes_idClientes, A.[cod_bonus] , A.[data_geracao] , A.[data_validade] , A.[moeda] , B.[razao_social] ,A.[data_resgate]"
			session("sql") = sql					
			if request.form("submit") = "Exportar" then
				call exportarParaArquivo(sql)
			end if
		else
			'sql = sql & " GROUP BY A.numero_solicitacao, A.Clientes_idClientes, A.[cod_bonus] , A.[data_geracao] , A.[data_validade] , A.[moeda] , B.[razao_social] ,A.[data_resgate]"		
			if session("sql") <> "" then
				sql = session("sql")
			else
				'call setRequest()
				'sql = sql & getWhereSQL()			
			end if			
		end if

'Response.Write sql & "<hr>"

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
                html = html & "<td "&style&">"&arr(19,i)&"</td>"
                html = html & "<td "&style&">"&arr(20,i)&"</td>"
                html = html & "<td "&style&">"&arr(21,i)&"</td>"
                html = html & "<td "&style&">"&arr(22,i)&"</td>"
                html = html & "<td "&style&">"&arr(23,i)&"</td>"
                html = html & "<td "&style&">"&arr(24,i)&"</td>"
                html = html & "<td "&style&">"&arr(25,i)&"</td>"
                html = html & "<td "&style&">"&arr(26,i)&"</td>"
                html = html & "<td "&style&">"&arr(27,i)&"</td>"
                html = html & "<td "&style&">"&arr(28,i)&"</td>"
                html = html & "<td "&style&">"&arr(29,i)&"</td>"


				'if len(trim(arr(5,i))) > 0 then
				'	html = html & "<td "&style&">"&DateRight(formatdatetime(arr(5,i),2))&"</td>"
				'else
				'	html = html & "<td "&style&"></td>"
				'end if	
				html = html & "</tr>"
			Next
			
			html = html & "<tr><td colspan=10>"
			html = html & "</td></tr>"
		else
			html = html & "<tr><td colspan='10' align='center' class='classColorRelPar'><b>Nenhum Bônus encontrado</b></td></tr>"
		end if
		getBonusCliente = html	  
	end function
	
	sub setRequest()
		statusBonus = Trim(Request.Form("status"))
		razaoSocial = Request.Form("razaosocial")
		dataGeracao_de = Trim(Request.Form("dedatacadastro"))
		dataGeracao_ate = Trim(Request.Form("atedatacadastro"))
		dataExpiracao_de = Trim(Request.Form("dedatacadastro2"))
		dataExpiracao_ate = Trim(Request.Form("atedatacadastro2"))
		'dataResgate_de = Trim(Request.Form("dedatacadastro3"))
		'dataResgate_ate = Trim(Request.Form("atedatacadastro3"))
		descBonus = Request.Form("razaosocial2")
		ufPonto = Trim(Request.Form("uf"))
		
		
'		Response.Write statusBonus & "<br />"
'		Response.Write razaoSocial & "<br />"
'		Response.Write dataGeracao_de & "<br />"
'		Response.Write dataGeracao_ate & "<br />"
'		Response.Write dataExpiracao_de & "<br />"
'		Response.Write dataExpiracao_ate & "<br />"
'		Response.Write dataResgate_de & "<br />"
'		Response.Write dataResgate_ate & "<br />"
'		Response.Write descBonus & "<br />"
	end sub
	
	function existWhere()
		if len(razaoSocial) > 0 or _
			 len(dataGeracao_de) > 0 and len(dataGeracao_ate) > 0 or _
			 len(dataExpiracao_de) > 0 and len(dataExpiracao_ate) > 0 or _
			 len(dataResgate_de) > 0 and len(dataResgate_ate) > 0 or _
			 len(descBonus) > 0 then
			existWhere = true
		else
			existWhere = false
		end if	
	end function
	
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
			convertDataSQL = ano & "-" & mes & "-" & dia
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
	
	function getWhereSQL()
		dim sql
		dim bAnd
        bAnd = true
		sql  = ""
		if existWhere() then
			'sql = sql & " where "
			'if statusBonus <> "todos" then
			'	if statusBonus = "gerado" then
			'		sql = sql & " A.[data_geracao] <> ''"	
			'		bAnd = true
			'	elseif statusBonus = "expirado" then
			'		sql = sql & " A.[data_geracao] <> '' and A.[data_validade] < convert(datetime, '"&date()&"') "
			'		bAnd = true
			'	else
			'		sql = sql & " A.[data_geracao] <> '' and A.[data_validade] <> '' and A.data_resgate <> ''"
			'		bAnd = true
			'	end if	
			'end if		
			if len(razaoSocial) > 0 then
				if bAnd then
					sql = sql & " and ( upper(CLI.[razao_social]) like upper('%"&razaoSocial&"%') or upper(CLI.nome_fantasia) like upper('%"&razaoSocial&"%') )  "	
				else
					sql = sql & " ( upper(CLI.[razao_social]) like upper('%"&razaoSocial&"%') or upper(CLI.nome_fantasia) like upper('%"&razaoSocial&"%') )  "	
					bAnd = true	
				end if
			end if	
			if len(dataGeracao_de) > 0 and len(dataGeracao_ate) > 0 then
				if bAnd then
					sql = sql & " and (CAST(FLOOR(CAST(SOLCOL.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataGeracao_de) & "' and '" & convertDataSQL(dataGeracao_ate) & "')"
				else
					sql = sql & " (CAST(FLOOR(CAST(SOLCOL.[data_solicitacao] AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataGeracao_de) & "' and '" & convertDataSQL(dataGeracao_ate) & "')"
					bAnd = true
				end if
			end if 	
			if len(dataExpiracao_de) > 0 and len(dataExpiracao_ate) > 0 then
				if bAnd then
					'sql = sql & " and A.[data_validade] between convert(datetime, '" & convertDataSQL(dataExpiracao_de) & "') and  convert(datetime,'" & convertDataSQL(dataExpiracao_ate) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(SOLCOL.[data_validade] AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataExpiracao_de) & "' and '" & convertDataSQL(dataExpiracao_ate) & "')"
				else
					'sql = sql & " A.[data_validade] between convert(datetime, '" & convertDataSQL(dataExpiracao_de) & "') and  convert(datetime,'" & convertDataSQL(dataExpiracao_ate) & "')"
					sql = sql & " (CAST(FLOOR(CAST(SOLCOL.[data_validade] AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataExpiracao_de) & "' and '" & convertDataSQL(dataExpiracao_ate) & "')"
					bAnd = true
				end if
			end if 	
			if len(dataResgate_de) > 0 and len(dataResgate_ate) > 0 then
				if bAnd then
					'sql = sql & " and A.[data_resgate] between convert(datetime, '" & convertDataSQL(dataResgate_de) & "') and  convert(datetime,'" & convertDataSQL(dataResgate_ate) & "')"
					sql = sql & " and (CAST(FLOOR(CAST(SOLCOL.[data_solicitacao_resgate]  AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataResgate_de) & "' and '" & convertDataSQL(dataResgate_ate) & "')"
				else
					'sql = sql & " A.[data_resgate] between convert(datetime, '" & convertDataSQL(dataResgate_de) & "') and  convert(datetime,'" & convertDataSQL(dataResgate_ate) & "')"
					sql = sql & " (CAST(FLOOR(CAST(SOLCOL.[data_solicitacao_resgate]  AS float)) AS datetime) BETWEEN '" & convertDataSQL(dataResgate_de) & "' and '" & convertDataSQL(dataResgate_ate) & "')"
					bAnd = true
				end if
			end if 
			if len(descBonus) > 0 then
				if bAnd then
					sql = sql & " and C.[descricao] like '%"&descBonus&"%'"	
				else
					sql = sql & " C.[descricao] like '%"&descBonus&"%'"
					bAnd = true	
				end if
			end if	
			'if ufPonto <> "" then
			'	if bAnd then
			'		sql = sql & " and B.[estado] = '"&ufPonto&"'"	
			'	else
			'		sql = sql & " B.[estado] = '"&ufPonto&"'"
			'		bAnd = true	
			'	end if
			'end if	
		end if
		getWhereSQL = sql		
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
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
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
	</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container" style="width:100%;">
  <table cellspacing="0" cellpadding="0" width="100%">
    <form action="" name="form1" method="POST">
      <tr>
        <td id="conteudo"><table cellspacing="3" cellpadding="2" width="100%" border=0>
            <tr>
              <td colspan="3" id="explaintitle" align="center">Relatório de Bônus do Cliente</td>
            </tr>
            <tr>
              <td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a></td>
            </tr>
            <tr>
              <td colspan="3"><table cellpadding="1" cellspacing="1" width="100%">
                  <tr>
                    <td width="80%"><fieldset style="font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif;">
                      <legend style="color:#666666;font-weight:bold;">Filtros</legend>
                      <div align="left" style="padding:3px 3px 3px 3px;width:100%;"> 
												<!--
												Status:                        
                        <select name="status" class="select" style="width:200px;">
                          <option value="todos" <% If Trim(Request.Form("status")) = "todos" Then %> selected <% End If %>>Todos</option>
                          <option value="gerado" <% If Trim(Request.Form("status")) = "gerado" Then %> selected <% End If %>>Gerado</option>
                          <option value="resgatado" <% If Trim(Request.Form("status")) = "resgatado" Then %> selected <% End If %>>Resgatado</option>
                          <option value="expirado" <% If Trim(Request.Form("status")) = "expirado" Then %> selected <% End If %>>Expirado</option>
                        </select>
                        &nbsp;&nbsp;&nbsp; <br>
                        -->
                        Razão Social: 
                        <input name="razaosocial" type="text" class="text" value="<%=Request.Form("razaosocial")%>" size="170" />
                      </div>
                      <div align="left" style="padding:3px 3px 3px 3px;width:100%;">Data de Gera&ccedil;&atilde;o  -
                        De:
                        <input name="dedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata1" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro','pop1','150',document.form1.dedatacadastro.value)" />
                        <span id="pop1" style="position:absolute;margin-left:20px;"></span> Até:
                        <input name="atedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("atedatacadastro"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata2" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro','pop2','150',document.form1.atedatacadastro.value)" />
                        <span id="pop2" style="position:absolute;margin-left:20px;"></span> 
                      </div>
                      <div align="left" style="padding:3px 3px 3px 3px;width:100%;">
                        Data de Expira&ccedil;&atilde;o - De: 
                        <input name="dedatacadastro2" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro2"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata12" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro2','pop3','150',document.form1.dedatacadastro2.value)" />
                        <span id="pop3" style="position:absolute;margin-left:20px;"></span> At&eacute;:
                        <input name="atedatacadastro2" type="text" class="text" value="<%=Trim(Request.Form("atedatacadastro2"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata22" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro2','pop4','150',document.form1.atedatacadastro2.value)" />
                        <span id="pop4" style="position:absolute;margin-left:20px;"></span> 
                      </div>  

                        <!--
                        <br>
                        Data de Resgate - De: 
                        <input name="dedatacadastro3" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro3"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata13" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro3','pop5','150',document.form1.dedatacadastro3.value)" />
                        <span id="pop5" style="position:absolute;margin-left:20px;"></span> At&eacute;:
                        <input name="atedatacadastro3" type="text" class="text" value="<%=Trim(Request.Form("atedatacadastro3"))%>" size="13" readonly />
                        <input TYPE="button" NAME="btndata23" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro3','pop6','150',document.form1.atedatacadastro3.value)" />
                        <span id="pop6" style="position:absolute;margin-left:20px;"></span></div>
                        -->
                      <div align="left" style="padding:3px 3px 3px 3px;width:100%;">Desc. B&ocirc;nus :
                        <input name="razaosocial2" type="text" class="text" value="<%=Request.Form("razaosocial2")%>" size="170" />
                      </div>

                        <div align="left" style="padding:3px 3px 3px 3px;width:100%;">
                            <input type="submit" name="submit" class="btnform" value="Procurar" />
                            <input name="submit" type="submit" class="btnform" value="Exportar" />
                            &nbsp;<a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a>
                        </div>

                        <div align="right" style="padding:3px 3px 3px 3px;width:100%;">
											<%'if session("sql") <> "" then%>
												<!--<a href="frmrelatoriobonuscliente.asp?rm=1">Clique aqui para refazer a pesquisa</a>-->
                    	<%'end if%>						  
														<input name="submit" type="submit" class="btnform" value="Procurar" />
						  
						  <input name="submit" type="submit" class="btnform" value="Exportar" />
                        </div>
                      </div>
                      </fieldset></td>
                  </tr>
                </table></td>
            </tr>
            <tr>
              <td colspan="3"><table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente" style="border:1px solid #000000">
                  <tr>
                    <th>Número da Solicitação de Coleta	</th>
                    <th>Usuário Solicitante do Pedido de Coleta	</th>
                    <th>Descr.Categoria	</th>
                    <th>Razão Social	</th>
                    <th>Nome fantasia	</th>
                    <th>CNPJ / CPF	</th>
                    <th>IE	</th>
                    <th>CEP</th>
                    <th>Logradouro	</th>
                    <th>Complemento Logradouro	</th>
                    <th>n°	</th>
                    <th>Bairro	</th>
                    <th>Municipio	</th>
                    <th>Estado	</th>
                    <th>DDD	</th>
                    <th>Telefone	</th>
                    <th>Qtd. Consumíveis Inservíveis 	</th>
                    <th>Qtd. De Volumes Enviados	</th>
                    <th>Data da Emissão da Nota Fiscal de Coleta	</th>
                    <th>N° Nota Fiscal da Coleta	</th>
                    <th>Data da Coleta no Cliente	</th>
                    <th>Data de chegada no Armazém	</th>
                    <th>Qtd. Inservíveis Recebidos / Conferidos	</th>
                    <th>Qtd. De Volumes Recebidos / Conferidos	</th>
                    <th>Códigos Recebidos	</th>
                    <th>Grupo de Produto	</th>
                    <th>Descrição	</th>
                    <th>Pontuação Unitária	</th>
                    <th>Pontuação Total	</th>
                    <th>Data da Geração do Bônus	</th>
                    <th>Status</th>

                    <!--<th>Data Resgate</th>-->
                  </tr>
				  <%
'				  if request("pag") = "" then
				  response.write getBonusCliente()
'				  end if
				  %>
                </table></td>
            </tr>
          </table></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
<%Call close()%>
