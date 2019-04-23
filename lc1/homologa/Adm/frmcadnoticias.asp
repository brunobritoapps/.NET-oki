<!--#include file="../_config/_config.asp" -->

<%Call open()%>
<%Call GetSessionAdm()%>
<%
	dim oFCKeditor
	dim texto
	dim ativo
	dim id
	dim fonte
	dim data
	dim titulo
	
	'set oFCKeditor = New FCKeditor

	function getText()
	'oFCKeditor.BasePath	= "../fckeditor/"
	'oFCKeditor.Create "editor1"
		response.write texto
	end function
	
	sub insere()
		dim sql
		sql = "INSERT INTO [marketingoki2].[dbo].[Home_Noticias] " & _
					   "([titulo] " & _
					   ",[text] " & _
					   ",[data] " & _
					   ",[fonte] " & _
					   ",[ativo]) " & _
				 "VALUES " & _
					   "('"&request.form("titulo")&"' " & _
					   ",'"&request.form("editor1")&"' " & _
					   ",convert(datetime, '"&FormatDate(request.form("data1"))&"') " & _
					   ",'"&request.form("fonte")&"' " & _
					   ","&cint(request.form("cbativo"))&")"
		call exec(sql)			   
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

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano
		
		Dia = day(sData)
		if Dia < 10 then Dia = "0" & Dia

		Mes = month(sData)
		if Mes < 10 then Mes = "0" & Mes

		Ano = year(sData)

		DateRight = Dia & "/" & Mes & "/" & Ano
	End Function
	
	function getNoticiaById(id)
		dim sql, arr, intarr, i
		sql = "select * from home_noticias where idnoticia = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				titulo = arr(1,i)
				texto = arr(2,i)
				data = DateRight(arr(3,i))
				fonte = arr(4,i)
				ativo = arr(5,i)
			next
		end if
	end function
	
	function getNoticias()
		dim sql, arr, intarr, i
		dim html, style
		html = ""		
		sql = "select * from home_noticias"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then	
					style = "class=""classColorRelPar""" 
				else
					style = "class=""classColorRelImpar""" 
				end if
				html = html & "<tr>"
				html = html & "<td "&style&"><img src=""img/buscar.gif"" class=""imgexpandeinfo"" onclick=""window.location.href='frmcadnoticias.asp?id="&arr(0,i)&"'"" /></td>"
				html = html & "<td "&style&">"&arr(1,i)&"</td>"
				html = html & "<td "&style&">"&arr(4,i)&"</td>"
				html = html & "<td "&style&">"&DateRight(arr(3,i))&"</td>"
				if arr(5,i) = 1 then
				html = html & "<td "&style&">Sim</td>"
				else
				html = html & "<td "&style&">Não</td>"
				end if
				html = html & "</tr>"
			next
		else
			html = html & "<tr><td colspan=""5"" align=""center"">Nenhum registro encontrado</td></tr>"
		end if
		getNoticias = html
	end function
	
	sub updateTexto(id)
		dim sql
		sql = "UPDATE [marketingoki2].[dbo].[Home_Noticias] " & _
				   "SET [titulo] = '"&request.form("titulo")&"' " & _
					  ",[text] = '"&request.form("editor1")&"' " & _
					  ",[data] = convert(datetime, '"&FormatDate(request.form("data1"))&"') " & _
					  ",[fonte] = '"&request.form("fonte")&"' " & _
					  ",[ativo] = "&request.form("cbativo")&" " & _
				 "WHERE idnoticia = " & id	 
		call exec(sql)	
		response.redirect "frmcadnoticias.asp"   
	end sub
	
	sub deletar(id)
		dim sql
		sql = "delete from home_noticias where idnoticia = " & id
		call exec(sql)
		response.redirect "frmcadnoticias.asp"   
	end sub
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		if request.form("action") = "Deletar" then
			call deletar(request.form("id"))
		elseif request.form("action") = "Editar" then
			call updateTexto(request.form("id"))
		else
			call insere()
		end if	
	else
		if len(trim(request.querystring("id"))) > 0 then
			id = request.querystring("id")
			call getNoticiaById(id)
			'oFCKeditor.Value = texto
		else
			'oFCKeditor.Value = ""
		end if	
	end if
%>
<html>
<head>
<script type="text/javascript" src="ckeditor/ckeditor.js"></script>
   <script type="text/javascript">
      window.onload = function()  {
        CKEDITOR.replace( 'editor1' );
      };
    </script> 
<SCRIPT LANGUAGE="JavaScript" SRC="js/CalendarPopup.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var cal = new CalendarPopup();
</SCRIPT>
<link rel="stylesheet" type="text/css" href="../css/geral.css">

<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="" name="form1" method="POST">
			<input type="hidden" name="id" value="<%=id%>" />
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<div id="painelcontrole">
						<table cellspacing="3" cellpadding="2" width="100%" border="0" id="tablelisttransportadoras">
							<tr>
								<td colspan="3" id="explaintitle" align="center">Cadastro de Notícias do Site</td>
							</tr>
							<tr>
								<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td colspan="3" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">Noticia:</td>
							</tr>	
							<tr>
								<td colspan="3"><textarea id="editor1" name="editor1" value="<%=texto%>"><%=getText()%></textarea></td>
							</tr>
							<tr>
								<td align="right" width="5%" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">Título: </td>
								<td align="left" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;"> <input type="text" name="titulo" value="<%=titulo%>" class="text" size="40" /></td>
							</tr>
							<tr>
								<td align="right" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">Fonte: </td>
								<td align="left" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;"> <input type="text" name="fonte" value="<%=fonte%>" class="text" size="40" /></td>
							</tr>
							<tr>
								<td align="right" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">Data:&nbsp; </td>
								<td align="left" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;"> <input type="text" name="data1" class="text" value="<%=data%>" /><A HREF="#" onClick="cal.select(document.forms['form1'].data1,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img align="absmiddle" src="img/btn_calendario.gif" border="0"></A></td>
							</tr>
							<tr>
								<td align="right" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">
									Ativo:&nbsp;
								</td>
								<td align="left" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">
									<select name="cbativo" class="select">
										<option value="">[Selecione]</option>
										<option value="0" <%if ativo = 0 then%>selected<%end if%>>Não</option>
										<option value="1" <%if ativo = 1 then%>selected<%end if%>>Sim</option>
									</select></td>
							</tr>
							<tr>
								<%if len(trim(request.querystring("id"))) > 0 then%>
									<td align="right" colspan="2">
										<input type="submit" name="action" class="btnform" value="Editar" />
										<input type="submit" name="action" class="btnform" value="Deletar" />
									</td>
								<%else%>	
									<td align="right"><input type="submit" name="action" class="btnform" value="Salvar"/></td>
								<%end if%>
							</tr>
							<tr>
								<td align="center" width="100%" colspan="2">
									<table width="100%" cellspacing="1" cellpadding="1" id="tableRelSolPendente" style="border:1px solid #000000">
										<tr>
											<th width="2%"><img src="img/check.gif" /></th>
											<th>Titulo</th>
											<th width="25%">Fonte</th>
											<th width="20%">Data</th>
											<th width="10%">Ativo</th>
										</tr>
										<%=getNoticias()%>
									</table>
								</td>
							</tr>
						</table>
					</div>
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
