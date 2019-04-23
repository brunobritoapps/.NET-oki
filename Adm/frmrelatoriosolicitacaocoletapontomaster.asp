<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
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
		dim i, arr, intarr
		dim arquivo
		dim fso
		dim arquivoPath
		dim filenamecsv
		dim filename
		dim cabecalhoArq
		
		set fso = server.createobject("scripting.filesystemobject")
		filenamecsv = "exportacao_relatorio_cliente_"&day(now())&"-"&month(now())&"-"&year(now())&"-"&fix(timer())&".csv"
		filename = request.servervariables("APPL_PHYSICAL_PATH") & "adm/exportacao/"&filenamecsv
		set arquivoPath = fso.createtextfile(filename)
		arquivo = ""
		call search(sql, arr, intarr)
		if intarr > -1 then
			cabecalhoArq = "Data Solicitação;Número Solicitação;Cod. Ponto Coleta;Razão Social Ponto Coleta;Qtd. Cartuchos;Data Programada para Coleta" 
			arquivoPath.writeLine(cabecalhoArq)
			for i=0 to intarr
				'arquivo = DateRight(formatdatetime(arr(5,i),2))&";"&arr(2,i)&";"&arr(15,i)&";"&arr(16,i)&";"&arr(3,i)&";"&DateRight(formatdatetime(arr(7,i)))
				arquivo = DateRight(arr(5,i))&";"&arr(2,i)&";"&arr(15,i)&";"&arr(16,i)&";"&arr(3,i)&";"&DateRight(arr(7,i))
				arquivoPath.writeLine(arquivo)
			next
		end if
		if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
			response.Redirect "http://www.sustentabilidadeoki.com.br/adm/exportacao/"&filenamecsv
		else
			response.Redirect "http://localhost:81/sgrs/adm/exportacao/"&filenamecsv
		end if
	end sub
	
	function getSolicitacaoByPonto()
		dim sql, arr, intarr, i
		dim html, style
		
		sql = "SELECT A.[idSolicitacao_coleta] " & _
				  ",A.[Status_coleta_idStatus_coleta] " & _
				  ",A.[numero_solicitacao_coleta] " & _
				  ",A.[qtd_cartuchos] " & _
				  ",A.[qtd_cartuchos_recebidos] " & _
				  ",A.[data_solicitacao] " & _
				  ",A.[data_aprovacao] " & _
				  ",A.[data_programada] " & _
				  ",A.[data_envio_transportadora] " & _
				  ",A.[data_entrega_pontocoleta] " & _
				  ",A.[data_recebimento] " & _
				  ",A.[motivo_status] " & _
				  ",A.[isMaster] " & _
				  ",B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
				  ",B.[Pontos_coleta_idPontos_coleta] " & _
				  ",C.[idPontos_coleta] " & _
				  ",C.[razao_social] " & _
				  ",C.[nome_fantasia] " & _
				  ",C.[cnpj] " & _
				  ",C.[bonus_type] " & _
				  ",C.[logradouro] " & _
				  ",C.[numero_endereco] " & _
				  ",C.[complemento_endereco] " & _
				  ",C.[bairro] " & _
				  ",C.[ddd] " & _
				  ",C.[telefone] " & _
				  ",C.[cep] " & _
				  ",C.[municipio] " & _
				  ",C.[estado] " & _
				  ",C.[usuario] " & _
				  ",C.[senha] " & _
				  ",C.[status_pontocoleta] " & _
				  ",C.[Qtd_Limite_Cartuchos] " & _
				  ",C.[idtransp] " & _
			  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_Pontos_coleta] AS B " & _
			  "ON A.[idSolicitacao_coleta] = B.[Solicitacao_coleta_idSolicitacao_coleta] " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Pontos_coleta] AS C " & _
			  "ON B.[Pontos_coleta_idPontos_coleta] = C.[idPontos_coleta] " & _
			  "where B.[Pontos_coleta_idPontos_coleta] = " & session("IDPonto") & " and left(A.[numero_solicitacao_coleta],1) = 'M'"

		if request.ServerVariables("HTTP_METHOD") = "POST" then
			call setRequest()
			if existWherePonto() then
				sql= sql & getWhereSQLPonto()
			end if
			if request.form("submit") = "Exportar" then
				call exportarParaArquivo(sql)
			end if
		end if
		'response.write "sql ponto: " & sql & "<br />"
		'response.end
			  
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				html = html & "<tr>"
				if not isnull(arr(5,i)) then
					html = html & "<td "&style&">"&DateRight(formatdatetime(arr(5,i),2))&"</td>"
				else
					html = html & "<td "&style&"></td>"
				end if	
				html = html & "<td "&style&">"&arr(2,i)&"</td>"
				html = html & "<td "&style&">"&arr(15,i)&"</td>"
				html = html & "<td "&style&">"&arr(16,i)&"</td>"
				html = html & "<td "&style&">"&arr(3,i)&"</td>"
				if not isnull(arr(7,i)) then
					html = html & "<td "&style&">"&DateRight(formatdatetime(arr(7,i),2))&"</td>"
				else
					html = html & "<td "&style&"></td>"
				end if	
				html = html & "</tr>"
			next
		else
			html = html & "<tr>"
			html = html & "<td colspan=""9"" align=""center"" class=""classColorRelPar""><b>Nenhum registro encontrado</b></td>"
			html = html & "</tr>"
		end if
		getSolicitacaoByPonto = html
	end function
	
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
	
	sub setRequest()
		tipoSolicitacao = Trim(Request.Form("tipo"))
		statusSolicitacao = Trim(Request.Form("status"))
		razaoSocial = Request.Form("razaosocial")
		transportadora = Trim(Request.Form("transportadora"))
		pontoColeta = Trim(Request.Form("pontocoleta"))
		ufCliente = Trim(Request.Form("uf"))
		dataSolicitacao_de = Trim(Request.Form("dedatacadastro"))
		dataSolicitacao_ate = Trim(Request.Form("atedatacadastro"))
		dataAprovacao_de = Trim(Request.Form("dedataaprovacao"))
		dataAprovacao_ate = Trim(Request.Form("atedataaprovacao"))
		dataProgramada_de = Trim(Request.Form("dedataprogramada"))
		dataProgramada_ate = Trim(Request.Form("atedataprogramada"))
		dataRecebimento_de = Trim(Request.Form("dedatarecebimento"))
		dataRecebimento_ate = Trim(Request.Form("atedatarecebimento"))
		dataEntregaPonto_de = Trim(Request.Form("dedataentrega"))
		dataEntregaPonto_ate = Trim(Request.Form("atedataentrega"))
		
		
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
	
	function existWherePonto()
		if len(Trim(Request.Form("dedatacadastro"))) > 0 or _
			 len(Trim(Request.Form("atedatacadastro"))) > 0 or _
			 len(Trim(Request.Form("dedataentrega"))) > 0 or _
			 len(Trim(Request.Form("atedataentrega"))) > 0 then
			existWherePonto = true
		else
			existWherePonto = false
		end if	
	end function

	function getWhereSQLPonto()
		dim sql
		dim bAnd
		bAnd = false
		if existWherePonto() then
			sql = sql & " and "
			if len(Trim(Request.Form("dedatacadastro"))) > 0 and len(Trim(Request.Form("atedatacadastro"))) > 0 then
				if bAnd then
					sql = sql & " and A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
				else
					sql = sql & " A.[data_solicitacao] between convert(datetime, '" & convertDataSQL(Request.Form("dedatacadastro")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedatacadastro")) & "')"
					bAnd = true
				end if
			end if
			if len(Trim(Request.Form("dedataentrega"))) > 0 and len(Trim(Request.Form("atedataentrega"))) > 0 then
				if bAnd then
					sql = sql & " and A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataentrega")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataentrega")) & "')"
				else
					sql = sql & " A.[data_programada] between convert(datetime, '" & convertDataSQL(Request.Form("dedataentrega")) & "') and  convert(datetime,'" & convertDataSQL(Request.Form("atedataentrega")) & "')"
					bAnd = true
				end if
			end if
		else
			sql = ""
		end if	
		getWhereSQLPonto = sql
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
			<td id="conteudo">
				<table cellspacing="3" cellpadding="2" width="100%" border=0>
					<tr>
						<td colspan="3" id="explaintitle" align="center">Relatório de Solicitação de Coleta</td>
					</tr>
					<tr>
						<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmtiporelatorioponto.asp';">&laquo Voltar</a></td>
					</tr>
					<tr>
						<td colspan="3">
							<table cellpadding="1" cellspacing="1" width="100%">
								<tr>
									<td width="80%">
										<fieldset style="font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif;">
											<legend style="color:#666666;font-weight:bold;">Filtros</legend>
										    <div align="left" style="padding:3px 3px 3px 3px;width:100%;">
												Data da Solicitação -
												De: 
												<input name="dedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("dedatacadastro"))%>" size="13" /> 
												<input TYPE="button" NAME="btndata1" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedatacadastro','pop1','150',document.form1.dedatacadastro.value)" /><span id="pop1" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedatacadastro" type="text" class="text" value="<%=Trim(Request.Form("atedatacadastro"))%>" size="13"  /> 
												<input TYPE="button" NAME="btndata2" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedatacadastro','pop2','150',document.form1.atedatacadastro.value)" /><span id="pop2" style="position:absolute;margin-left:20px;"></span>
										</div>
											<div align="left" style="padding:3px 3px 3px 3px;width:100%;">Data Programada para a Coleta -
												De: 
											  <input name="dedataentrega" type="text" class="text" value="<%=Trim(Request.Form("dedataentrega"))%>" size="13" /> 
												<input TYPE="button" NAME="btndata9" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.dedataentrega','pop9','150',document.form1.dedataentrega.value)" /><span id="pop9" style="position:absolute;margin-left:20px;"></span>
												Até: <input name="atedataentrega" type="text" class="text" value="<%=Trim(Request.Form("atedataentrega"))%>" size="13"  /> 
												<input TYPE="button" NAME="btndata0" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.atedataentrega','pop0','150',document.form1.atedataentrega.value)" /><span id="pop0" style="position:absolute;margin-left:20px;"></span>											</div>
											<div align="right" style="padding:3px 3px 3px 3px;width:100%;">
												<input type="submit" class="btnform" name="submit" value="Procurar" />
												<input name="submit" type="submit" class="btnform" value="Exportar" />
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
								<% If not cint(Request.Form("tipo")) = 1 Then %>
								<tr>
									<th colspan="8" id="explaintitle">Ponto de Coleta </th>
								</tr>
								<tr>
									<th>Data Solicitação</th>
									<th>Número Solicitação</th>
									<th>Cod. Ponto Coleta </th>
									<th>Razão Social Ponto Coleta</th>
									<th>Qtd. Cartuchos</th>
									<th>Data Programada para Coleta</th>
								</tr>
								<%= getSolicitacaoByPonto() %>
								<% End If %>
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
