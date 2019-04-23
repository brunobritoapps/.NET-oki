<!--#include file="../_config/_config.asp" -->
<%
if request("rm") = "1" then
session("sql") = ""
response.redirect("frmrelatoriocoletaaprovada.asp")
end if
%>

<%Call open()%>
<%Call GetSessionAdm()%>
<%
    dim sfiltra 'peterson 17-5-2014
	dim razaosocial
	dim status_cliente
	dim grupo
	dim categoria
	dim cod_bonus_de
	dim cod_bonus_ate
	dim data_cadastro_de
	dim data_cadastro_ate

    Sub exportarParaArquivo(sql)
                'refeito peterson aquino: 17-5-2014
				'call exportarParaArquivo(sql)
    		    'tipoSolicitacao = Trim(Request.Form("tipo"))
		        'statusSolicitacao = Trim(Request.Form("status"))
		        'dataSolicitacao_de = convertDataSQL(Request.Form("dedatacadastro"))
		        'dataSolicitacao_ate = convertDataSQL(Request.Form("atedatacadastro")) //convertDataSQL(Request.Form("dedatacadastro"))
                response.Redirect "rptcoletasaprovadas.aspx?filtra=" & sFiltra

    End Sub
	
	sub exportarParaArquivoOld(sql)

		dim i, arr, intarr
		dim arquivo
		dim fso
		dim arquivoPath
		dim filenamecsv
		dim filename
		dim cabecalhoArq
		
		set fso = server.createobject("scripting.filesystemobject")
		filenamecsv = "exportacao_relatorio_cliente_"&day(now())&"-"&month(now())&"-"&year(now())&"-"&fix(timer())&".csv"
        '		filename = request.servervariables("APPL_PHYSICAL_PATH") & "adm\exportacao\"&filenamecsv
        'Alteraçào feita por Wellington
        'Programador: Wellington
        'Descrição: Alterado o caminho para a gravação do arquivo
		filename = server.MapPath("exportacao\"&filenamecsv)
        '		response.write filename
        '		response.End()
		set arquivoPath = fso.createtextfile(filename)
		arquivo = ""
		call search(sql, arr, intarr)
		if intarr > -1 then
			cabecalhoArq = "Cod. Cliente;Razão Social;Cod. Categoria;Desc. Categoria;Cod. Grupo;Desc. Grupo;Cod. Bônus;Desc. Bônus;Data Cadastro;Status"
			arquivoPath.writeLine(cabecalhoArq)
			for i=0 to intarr
				arquivo = arr(0,i)&";"&arr(1,i)&";"&arr(19,i)&";"&arr(20,i)&";"&arr(21,i)&";"&arr(22,i)&";"&arr(23,i)&";"&arr(24,i)&";"&getDataCadastro(arr(0,i))&";"&GetStatusDesc(arr(25,i))
				arquivoPath.writeLine(arquivo)
			next
		end if
		if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
			response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/adm/exportacao/"&filenamecsv
		else
            response.Redirect "http://www.sustentabilidadeoki.com.br/lc/homologa/adm/exportacao/"&filenamecsv
			'response.Redirect "http://localhost:81/sgrs/adm/exportacao/"&filenamecsv
		end if
	end sub
	
	Sub getColetasAprovadas()
		Dim sSql, arrClientes, intClientes, i

        sSql = sSql & "select "
        sSql = sSql & "a.idSolicitacao_coleta "
        sSql = sSql & ",a.numero_solicitacao_coleta "
        sSql = sSql & ",a.data_aprovacao "
        sSql = sSql & ",c.Clientes_idClientes, c.cep_coleta, c.logradouro_coleta "
        sSql = sSql & ",c.numero_endereco_coleta, c.comp_endereco_coleta, c.bairro_coleta, c.municipio_coleta, c.estado_coleta "
        sSql = sSql & ",c.contato_coleta, c.ddd_resp_coleta, c.telefone_resp_coleta, c.depto_resp_coleta, c.ramal_resp_coleta "
        sSql = sSql & ",e.Transportadoras_idTransportadoras ,f.email ,e.razao_social, e.cnpj "
        sSql = sSql & ",c.ramal_resp_coleta "
        sSql = sSql & "from dbo.Solicitacao_coleta as a "
        sSql = sSql & "left outer join solicitacao_coleta_has_clientes as c on c.Solicitacao_coleta_idSolicitacao_coleta = a.idSolicitacao_coleta  "
        sSql = sSql & "left outer join Solicitacao_coleta as d on a.idSolicitacao_coleta = d.idSolicitacao_coleta  "
        sSql = sSql & "left outer join Clientes as e on e.idClientes = c.Clientes_idClientes  "
        sSql = sSql & "left outer join Transportadoras as f on f.idTransportadoras = e.Transportadoras_idTransportadoras  "
        sSql = sSql & " where d.Status_coleta_idStatus_coleta = 2 "

		if request.ServerVariables("HTTP_METHOD") = "POST" then	
				getRequest()
                sFiltra = getWhere()
				sSql = sSql & sFiltra

				session("sql") = sSql
			
				if request.form("submit") = "Exportar" then
					call exportarParaArquivo(sSql)
				end if

			else
			if session("sql") <> "" then
				sSql = session("sql")
			end if
		end if	

        'concatena a ordenação mesmo sem fltrar
        sSql = sSql & " order by a.idSolicitacao_coleta "
		
		'Response.Write ssql & "<hr>"
		
		Call search(sSql, arrClientes, intClientes)

        'Response.Write "<script>alert('" & intClientes & "')</script>"

		If intClientes > -1 Then 
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 50 'numero de registros mostrados na pagina 'peterson - 13-5-2014 aumentei de 30 para 50.
			intNumProds = intClientes+1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao

			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			Response.Write "<tr><td colspan=9><div id=pag>"
			Response.Write PaginacaoExibir(intPag, intProdsPorPag, intClientes)
			Response.Write "</div></td></tr>"
    
            'Response.Write "<script>alert('" & intClientes & "," & intUltima & "," & intPag & "')</script>"

			For i = (intProdsPorPag * (intPag - 1)) to intUltima			
				'if getDataCadastro(arrClientes(0,i)) <> "" then
                if arrClientes(0,i) <> "" then
					If i Mod 2 = 0 Then
						Response.Write "<tr>" 				

						Response.Write "<td class='classColorRelPar'>"&arrClientes(00,i)&"</td>"				'id
						Response.Write "<td class='classColorRelPar'>"&arrClientes(01,i)&"</td>"				'numero

						Response.Write "<td class='classColorRelPar' align='center' width='3%'>"&arrClientes(3,i)&"</td>"	'id cliente
						Response.Write "<td class='classColorRelPar'>"&arrClientes(18,i)&"</td>"				'razão social

                        Response.Write "<td class='classColorRelPar width='8%'>"&arrClientes(4,i)&"</td>"				'cep
						Response.Write "<td class='classColorRelPar'>"&arrClientes(05,i)&"</td>"				'endereço
						Response.Write "<td class='classColorRelPar'>"&arrClientes(07,i)&"</td>"				'complemento
						Response.Write "<td class='classColorRelPar'>"&arrClientes(08,i)&"</td>"				'bairro
						Response.Write "<td class='classColorRelPar'>"&arrClientes(09,i)&"</td>"				'Municpio
						Response.Write "<td class='classColorRelPar'>"&arrClientes(10,i)&"</td>"				'estado
						Response.Write "<td class='classColorRelPar'>"&arrClientes(11,i)&"</td>"				'contato
						Response.Write "<td class='classColorRelPar'>("&arrClientes(12,i)&") " & arrClientes(13,i)  & "</td>"				'telefone
						Response.Write "<td class='classColorRelPar'>"&arrClientes(15,i)&"</td>"				'ramal


						Response.Write "<td class='classColorRelPar'>"&arrClientes(14,i)&"</td>"				'departamento
						'Response.Write "<td class='classColorRelPar'>"&arrClientes(21,i)&"</td>"				
						'Response.Write "<td class='classColorRelPar'>"&arrClientes(22,i)&"</td>"				
						'Response.Write "<td class='classColorRelPar'>"&arrClientes(23,i)&"</td>"				
						'Response.Write "<td class='classColorRelPar'>"&arrClientes(24,i)&"</td>"				
						'Response.Write "<td class='classColorRelPar'>"&getDataCadastro(arrClientes(0,i))&"</td>"
                        'Response.Write "<td class='classColorRelPar'>"&right((day(arrClientes(26,i))+100),2)&"/"&right(((month(arrClientes(26,i)))+100),2)&"/"&year(arrClientes(26,i))&"</td>"
						'Response.Write "<td class='classColorRelPar'>"&GetStatusDesc(arrClientes(25,i))&"</td>"
						Response.Write "</tr>"
					Else
						Response.Write "<tr>" 				

						Response.Write "<td class='classColorRelImpar'>"&arrClientes(00,i)&"</td>"				'id
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(01,i)&"</td>"				'numero
	
    					Response.Write "<td class='classColorRelImpar' align='center' width='3%'>"&arrClientes(3,i)&"</td>"	'id cliente
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(18,i)&"</td>"				'razão social

                        Response.Write "<td class='classColorRelImpar width='8%'>"&arrClientes(4,i)&"</td>"				'cep
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(05,i)&"</td>"				'endereço
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(07,i)&"</td>"				'complemento
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(08,i)&"</td>"				'bairro
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(09,i)&"</td>"				'Municpio
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(10,i)&"</td>"				'estado
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(11,i)&"</td>"				'contato
						Response.Write "<td class='classColorRelImpar'>("&arrClientes(12,i)&") " & arrClientes(13,i)  & "</td>"				'telefone
						Response.Write "<td class='classColorRelImpar'>"&arrClientes(15,i)&"</td>"				'ramal


						Response.Write "<td class='classColorRelImpar'>"&arrClientes(14,i)&"</td>"				'departamento
						'Response.Write "<td class='classColorRelImpar'>"&arrClientes(21,i)&"</td>"				
						'Response.Write "<td class='classColorRelImpar'>"&arrClientes(22,i)&"</td>"				
						'Response.Write "<td class='classColorRelImpar'>"&arrClientes(23,i)&"</td>"				
						'Response.Write "<td class='classColorRelImpar'>"&arrClientes(24,i)&"</td>"				
						'Response.Write "<td class='classColorRelImpar'>"&getDataCadastro(arrClientes(0,i))&"</td>"
                        'Response.Write "<td class='classColorRelImpar'>"&right((day(arrClientes(26,i))+100),2)&"/"&right(((month(arrClientes(26,i)))+100),2)&"/"&year(arrClientes(26,i))&"</td>"
						'Response.Write "<td class='classColorRelImpar'>"&GetStatusDesc(arrClientes(25,i))&"</td>"
						Response.Write "</tr>"
					End If	
				else
                    'Alteraçào feita por Wellington
                    'Programador: Wellington
                    'Descrição: Mostrar apenas a mensagem quando não houver registros
				if i <=0 then
					Response.Write "<tr><td colspan='9' align='center' class='classColorRelPar'><b>Nenhum Cliente encontrado</b></td></tr>"
					end if
					exit for 
				end if	
			Next
			
			Response.Write "<tr><td colspan=9>"
            'Alteraçào feita por Wellington
            'Programador: Wellington
            'Descrição: passar deste if caso existe cadstro de cliente nesta data
			Response.Write "</td></tr>"
		Else
            Response.Write "<tr><td colspan='9' align='center' class='classColorRelPar'><b>Nenhum Cliente encontrado</b></td></tr>"
		End If						
	End Sub
	
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
	
	function getGrupos()
		dim sql, arr, intarr, i
		dim html
		dim selected
		sql = "SELECT [idGrupos] " & _
				  ",[descricao] " & _
			  "FROM [marketingoki2].[dbo].[Grupos]"
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			for i=0 to intarr
				if cint(request.form("grupos")) = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(1,i)&"</option>"
			next
		else
			html = html & "<option value="""">---</option>"	
		end if
		getGrupos = html
	end function
	
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
		DateRight = Dia & "/" & Mes & "/" & Ano
	End Function
	
	function getDataCadastro(idcliente)
		dim sql, arr, intarr, i
		sql = "select top 1 a.data_solicitacao from solicitacao_coleta as a " & _
				"left join solicitacao_coleta_has_clientes as b " & _
				"on a.idsolicitacao_coleta = b.Solicitacao_coleta_idSolicitacao_coleta " & _
				"where b.clientes_idclientes = " & idcliente
				
		if request.servervariables("HTTP_METHOD") = "POST" then
			sql = sql & getWhereDataCadastro()
		end if
'		response.write sql & "---<br><br><br>"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getDataCadastro = DateRight(formatdatetime(arr(0,i),2))
			next
		else
			getDataCadastro = ""
		end if		
	end function
	
	function getCategoria()
		dim sql, arr, intarr, i
		dim html
		dim selected
		sql = "SELECT [idCategorias] " & _
				  ",[descricao] " & _
				  ",[ativo] " & _
				  ",[isColetaDomiciliar] " & _
				  ",[minCartuchos] " & _
			  "FROM [marketingoki2].[dbo].[Categorias]"
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			for i=0 to intarr
				if cint(request.form("categoria")) = arr(0,i) then
					selected = "selected"
				else
					selected = ""
				end if
				html = html & "<option value="""&arr(0,i)&""" "&selected&">"&arr(1,i)&"</option>"
			next
		else
			html = html & "<option value="""">---</option>"	
		end if
		getCategoria = html
	end function
	
	function getRequest()
		data_cadastro_de = Trim(Request.Form("dedatacadastro"))
		data_cadastro_ate = Trim(Request.Form("atedatacadastro"))
		
'		Response.Write razaosocial & "<br />"
'		Response.Write status_cliente & "<br />"
'		Response.Write grupo & "<br />"
'		Response.Write categoria & "<br />"
'		Response.Write cod_bonus_de & "<br />"
'		Response.Write cod_bonus_ate & "<br />"
'		Response.Write data_cadastro_de & "<br />"
'		Response.Write data_cadastro_ate & "<br />"
'		response.end
	end function
	
	function existWhere()
		if len(trim(data_cadastro_de)) > 0 or _
			len(trim(data_cadastro_ate)) > 0 then
			'Alteraçào feita por peterson loop (11) 996140516 peterson.aquino@hotmail.com
			existWhere = true
		else
			existWhere = false
		end if	
	end function
	
	function getWhereDataCadastro()
		dim sql
		sql = ""
		if len(trim(data_cadastro_de)) > 0 and len(trim(data_cadastro_ate)) > 0 then
			'sql = sql & " and a.data_solicitacao between convert(datetime, '" & convertDataSQL(trim(data_cadastro_de)) & "') and  convert(datetime,'" & convertDataSQL(trim(data_cadastro_ate)) & "')"
			sql = sql & " and (CAST(FLOOR(CAST(a.data_solicitacao AS float)) AS datetime) BETWEEN '" & convertDataSQL(trim(data_cadastro_de)) & "' and '" & convertDataSQL(trim(data_cadastro_ate)) & "')"
		end if
		getWhereDataCadastro = sql
	end function

    '
    'peterson 17-5-2014
    'loop996140516
	function getWhere()
		dim sql
		dim bAnd
		
		bAnd = false
		sql = ""
		
		if existWhere() then	

			sql = sql & "and "

			if len(trim(data_cadastro_de)) > 0 and len(trim(data_cadastro_ate)) > 0 then
				if not bAnd then
					sql = sql & " (CAST(FLOOR(CAST(a.data_aprovacao AS float)) AS datetime) BETWEEN '" & convertDataSQL(trim(data_cadastro_de)) & "' and '" & convertDataSQL(trim(data_cadastro_ate)) & "')"
					bAnd = true
				else
					sql = sql & " and (CAST(FLOOR(CAST(a.data_aprovacao AS float)) AS datetime) BETWEEN '" & convertDataSQL(trim(data_cadastro_de)) & "' and '" & convertDataSQL(trim(data_cadastro_ate)) & "')"
				end if
			end if				
		else
			sql = ""
		end if
		getWhere = sql
	end function
	
	function GetStatusDesc(lStatus)
		select case lStatus
			case 0
				GetStatusDesc = "Status em branco"
			case 1
				GetStatusDesc = "Aprovado"
			case 2
				GetStatusDesc = "Rejeitado"
			case 3
				GetStatusDesc = "Inativo"
			case else
				GetStatusDesc = ""
		end select
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

function formValidate() {
	var windowForm = document.form1;
	if (windowForm.atebonus.value != "" && windowForm.debonus.value == "") {
		alert("Preencha o campo Cód de Bônus - De:");
		return;
	}
	windowForm.submit();
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
			<td id="conteudo">
				<table cellspacing="3" cellpadding="2" width="100%" border=0>
					<tr>
						<td colspan="3" id="explaintitle" align="center">Relatório de Coletas Aprovadas</td>
					</tr>
					<tr>
						<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a></td>
					</tr>
					<tr>
						<td colspan="3">
							<table cellpadding="1" cellspacing="1" width="100%">
								<tr>
									<td width="80%">
										<fieldset style="font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif;">
											<legend style="color:#666666;font-weight:bold;">Filtros</legend>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Data de Aprovação -
												De: <input type="text" class="text" name="dedatacadastro" value="<%= Request.Form("dedatacadastro") %>" size="13" readonly /> <input TYPE="button" NAME="btndata1" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro','pop1','150',document.form1.dedatacadastro.value)" /><span id="pop1" style="position:absolute;margin-left:20px;"></span>
												Até: <input type="text" class="text" name="atedatacadastro" value="<%= Request.Form("atedatacadastro") %>" size="13" readonly /> <input TYPE="button" NAME="btndata2" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro','pop2','150',document.form1.atedatacadastro.value)" /><span id="pop2" style="position:absolute;margin-left:20px; top: 63px; left: 383px; height: 17px;"></span>
											</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">
											<%'if session("sql") <> "" then%>
												<!--<a href="frmrelatoriocliente.asp?rm=1">Clique aqui para refazer a pesquisa</a>-->
                    	                        <%'end if%>						  
												<input type="submit" name="submit" class="btnform" value="Procurar" />
												<input name="submit" type="submit" class="btnform" value="Exportar" />
                                                &nbsp;<a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorio.asp';">&laquo Voltar</a>
											</div>
										</fieldset>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan="3">
							<table cellpadding="1" cellspacing="1" width="1700px" id="tableRelSolPendente" style="border:1px solid #000000">
								<tr>
                                    <th>Id</th>
                                    <th>Número</th>
									<th>Cod. Cliente</th>
									<th>Razão Social</th>
                                    <th>CEP</th>
                                    <th>Endereço</th>
                                    <th>Complemento</th>
                                    <th>Bairro</th>
                                    <th>Município</th>
                                    <th>Estado</th>
                                    <th>Contato</th>
                                    <th>Telefone</th>
                                    <th>Ramal</th>
									<th>Departamento</th>
								</tr>
								<%call getColetasAprovadas()%>
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
