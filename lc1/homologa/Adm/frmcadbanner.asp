<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	dim data_inicio
	dim data_termino
	dim link
	dim busca
	dim tipo
	dim fileimage
	dim id
	
	dim txt
	
	sub getBannerByImg(img)
		dim sql, arr, intarr, i
		sql = "SELECT day([data_inicio]) " & _
			  ",month([data_inicio]) " & _
			  ",year([data_inicio]) " & _
			  ",day([data_termino]) " & _
			  ",month([data_termino]) " & _
			  ",year([data_termino]) " & _
			  ",[link] " & _
			  ",[busca] " & _
			  ",[imagem] " & _
			  ",[tipo] " & _
			  "FROM [marketingoki2].[dbo].[Home_Banners] where [idbanner] = '"&img&"'"
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			for i=0 to intarr
				data_inicio = arr(0,i) & "/" & arr(1,i) & "/" & arr(2,i)
				data_termino = arr(3,i) & "/" & arr(4,i) & "/" & arr(5,i)
				link = arr(6,i)
				busca = arr(7,i)
				fileimage = arr(8,i)
				tipo = arr(9,i)
				id = img
			next
		else
			response.redirect "frmcadbanner.asp"
		end if
	end sub
	
	function getBanners()
		dim sql, arr, intarr, i
		dim html, style
		
		sql = "SELECT day([data_inicio]) " & _
				  ",month([data_inicio]) " & _
				  ",year([data_inicio]) " & _
				  ",day([data_termino]) " & _
				  ",month([data_termino]) " & _
				  ",year([data_termino]) " & _
				  ",[link] " & _
				  ",[busca] " & _
				  ",[imagem] " & _
				  ",[tipo] " & _
				  ",[idbanner] " & _
			 	"FROM [marketingoki2].[dbo].[Home_Banners]"
				
		call search(sql ,arr, intarr)	  
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if	
				html = html & "<tr>"
				html = html & "<td "&style&" width=""2%""><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Editar Banner"" onclick=""window.location.href='frmcadbanner.asp?id="&arr(10,i)&"'"" /></td>"
				html = html & "<td "&style&">"&arr(0,i)&"/"&arr(1,i)&"/"&arr(2,i)&"</td>"
				html = html & "<td "&style&">"&arr(3,i)&"/"&arr(4,i)&"/"&arr(5,i)&"</td>"
				html = html & "<td "&style&">"&arr(6,i)&"</td>"
				html = html & "<td "&style&">"&arr(7,i)&"</td>"
				html = html & "<td "&style&">"&arr(8,i)&"</td>"
				html = html & "<td "&style&">"&arr(9,i)&"</td>"
				html = html & "</tr>"
			next	
		else
			html = html & "<tr><td colspan=""6"" align=""center"">Nenhum registro encontrado</td></tr>"
		end if
		getBanners = html
	end function
	
	function getListFiles()
		dim files
		dim fso
		dim folder
		dim ofolder
		dim nome_file
		dim i
		dim nome
		dim selected

		i = 0
		folder = request.servervariables("APPL_PHYSICAL_PATH")&"adm\home\"
		set fso = server.createobject("scripting.filesystemobject")
		set ofolder = fso.getfolder(folder)
		set files = ofolder.files
		for each nome_file in files
			i = i + 1
			nome = split(nome_file, "\")
			
			txt = nome_file.Name
			
			'if fileimage = nome(6) then
			if fileimage = txt then
				selected = "selected"
			else
				selected = ""
			end if
			
			
			'getListFiles = getListFiles & "<option value="""&nome(6)&""" "&selected&">"&nome(5)&" / "&nome(6)&txt&"</option>"
			getListFiles = getListFiles & "<option value="""&txt&""" "&selected&">"&txt&"</option>"
		next
		
		set fso = nothing
		set ofolder = nothing
		set folder = nothing
	end function
	
	Sub Submit()
		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			data_inicio = request.form("data1")
			data_termino = request.form("data2")
			link = request.form("link")
			busca = request.form("busca")
			tipo = request.form("cbtipo")
			fileimage = request.form("cbimage")
			id = request.form("id")
			if len(trim(id)) > 0 then
				if request.form("action") = "Deletar" then
					call deletar(id)
				else
					call updateBanner(id)
				end if	
			else
				call insere()
			end if
			data_inicio = ""
			data_termino = ""
			link = ""
			busca = ""
			fileimage = ""
			tipo = ""
		else
			if len(trim(request.querystring("id"))) > 0 then
				call getBannerByImg(request.querystring("id"))	
			end if	
		End If
	End Sub

	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia
		dim dataSplit
		dataSplit = split(sDate, "/")
		Dia = trim(dataSplit(0))
		Mes = trim(dataSplit(1))
		Ano = trim(dataSplit(2))
		if len(Dia) = 0 then
			Dia = "0"&Dia
		end if
		if len(Mes) = 0 then
			Mes = "0"&Mes
		end if
		if len(Ano) = 0 then
			Ano = "0"&Ano
		end if
		FormatDate = Ano & "-" & Mes & "-" & Dia
	End Function
	
	sub updateBanner(id)
		dim sql
		sql = "UPDATE [marketingoki2].[dbo].[Home_Banners] " & _
			  "SET [data_inicio] = '"&FormatDate(data_inicio)&"' " & _
			  ",[data_termino] = '"&FormatDate(data_termino)&"' " & _
			  ",[link] = '"&link&"' " & _
			  ",[busca] = '"&busca&"' " & _
			  ",[imagem] = '"&fileimage&"' " & _
			  ",[tipo] = '"&tipo&"' " & _
			  "WHERE [idbanner] = " & id
		call exec(sql)
		response.redirect "frmcadbanner.asp"	  
	end sub
	
	sub deletar(id)
		dim sql
		sql = "delete from home_banners where idbanner = " & id
		call exec(sql)
		response.redirect "frmcadbanner.asp"	  
	end sub

	sub insere()
		dim sql
		sql = "INSERT INTO [marketingoki2].[dbo].[Home_Banners] " & _
					   "([data_inicio] " & _
					   ",[data_termino] " & _
					   ",[link] " & _
					   ",[busca] " & _
					   ",[imagem] " & _
					   ",[tipo]) " & _
				 "VALUES " & _
					   "('"&FormatDate(data_inicio)&"'" & _
					   ",'"&FormatDate(data_termino)&"'" & _
					   ",'"&link&"' " & _
					   ",'"&busca&"' " & _
					   ",'"&fileimage&"' " & _
					   ",'"&tipo&"')"
'		response.write sql
'		response.end			   
		call exec(sql)			   
	end sub
	
	call Submit()
%>
<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
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
  for (n=1991; n<2020; n++)
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

function validaFormBanner() {
	var form = document.form1;
	if (form.data1.value == "") {
		alert("Preencha o campo data de inicio.");
		return;
	}
	if (form.data2.value == "") {
		alert("Preencha o campo data de término.");	
		return;
	}
	if (form.cbimage.value == "") {
		alert("Escolha uma imagem para o Banner");
		return;
	}
	if (form.cbtipo.value == "") {
		alert("Escolha um tipo de Banner");
		return;
	}
	form.submit();
}

</script>
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
						<form action="" name="form1" method="POST">
						<input type="hidden" name="id" value="<%=id%>" />
						<table cellspacing="1" cellpadding="1" width="100%" id="tablelisttransportadoras">
							<tr>
								<td id="explaintitle" colspan="2" align="center">Cadastro do Banner</td>
							</tr>
							<tr>
								<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmoperacionaladm.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td align="right">Data inicio:</td>
								<td><input type="text" name="data1" size="10" class="textreadonly" value="<%=data_inicio%>" /> <input TYPE="button" NAME="btnData1" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.data1','pop1','150',document.form1.data1.value)" /> <span id="pop1" style="position:absolute;margin-left:110px;"></span></td>
							</tr>
							<tr>
								<td align="right">Data término:</td>
								<td><input type="text" name="data2" size="10" class="textreadonly" value="<%=data_termino%>" /> <input TYPE="button" NAME="btnData2" class="btnform" VALUE="..." Onclick="javascript:popdate('document.form1.data2','pop2','150',document.form1.data2.value)" /> <span id="pop2" style="position:absolute;margin-left:110px;"></span></td>
							</tr>
							<tr>
								<td align="right">Link do Banner:</td>
								<td><input type="text" name="link" class="textreadonly" value="<%=link%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right">Texto para Busca:</td>
								<td><input type="text" name="busca" class="textreadonly" value="<%=busca%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right">Imagem do Banner:</td>
								<td>
									<select name="cbimage" class="select">
										<option value="">[Selecioneii]</option>
										<%=getListFiles()%>
									</select>
								</td>
							</tr>
							<tr>
								<td align="right">Tipo:</td>
								<td>
									<select name="cbtipo" class="select">
										<option value="">[Selecione]</option>
										<option value="lateral" <%if tipo = "lateral" then%>selected<%end if%>>Lateral</option>
										<option value="rodape" <%if tipo = "rodape" then%>selected<%end if%>>Rodapé</option>
									</select>
								</td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<%if len(trim(request.querystring("id"))) > 0 then%>
									<td align="center" colspan="2">
										<input type="button" name="action" value="Salvar" class="btnform" onClick="validaFormBanner()" />&nbsp;
										<input type="submit" name="action" value="Deletar" class="btnform" />
									</td>
								<%else%>	
									<td align="center" colspan="2"><input type="button" name="action" value="Incluir" class="btnform" onClick="validaFormBanner()" /></td>
								<%end if%>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2">
									<table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente" style="border:1px solid #000000">
										<tr>
											<th><img src="img/check.gif" /></th>
											<th>Data início</th>
											<th>Data término</th>
											<th>Link</th>
											<th>Busca</th>
											<th>Imagem</th>
											<th>Tipo</th>
										</tr>
										<%=getBanners()%>
									</table>
								</td>
							</tr>
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
