<!--#include file="../_config/_config.asp" -->

<%Call open()%>
<%Call GetSessionAdm()%>
<%
	dim oFCKeditor
	dim texto
	dim area
	dim id
	
	'set oFCKeditor = New FCKeditor

	function getText()
		'oFCKeditor.BasePath	= "../fckeditor/"
		'oFCKeditor.Create "FCKeditor1"
        'response.write texto
	end function
	
	sub insere()
		dim sql
		sql = "INSERT INTO [marketingoki2].[dbo].[Home_Textos] " & _
					   "([texto] " & _
					   ",[area]) " & _
				 "VALUES " & _
					   "('"&request.form("editor1")&"' " & _
					   ",'"&request.form("cbarea")&"')"
		call exec(sql)			   
	end sub
	
	function getTextoById(id)
		dim sql, arr, intarr, i
		sql = "select * from home_textos where idtexto = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				texto = arr(1,i)
				area = arr(2,i)
			next
		end if
	end function
	
	function getTextos()
		dim sql, arr, intarr, i
		dim html, style
		html = ""		
		sql = "select * from home_textos"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then	
					style = "class=""classColorRelPar""" 
				else
					style = "class=""classColorRelImpar""" 
				end if
				html = html & "<tr>"
				html = html & "<td "&style&"><img src=""img/buscar.gif"" class=""imgexpandeinfo"" onclick=""window.location.href='frmcadtextos.asp?id="&arr(0,i)&"'"" /></td>"
				html = html & "<td "&style&">"&arr(1,i)&"</td>"
				html = html & "<td "&style&">"&arr(2,i)&"</td>"
				html = html & "</tr>"
			next
		else
			html = html & "<tr><td colspan=""3"" align=""center"">Nenhum registro encontrado</td></tr>"
		end if
		getTextos = html
	end function
	
	sub updateTexto(id)
		dim sql
		sql = "UPDATE [marketingoki2].[dbo].[Home_Textos] " & _
			   "SET [texto] = '"&request.form("editor1")&"' " & _
			   ",[area] = '"&request.form("cbarea")&"' " & _
			   "WHERE idtexto = " & id
		call exec(sql)	
		response.redirect "frmcadtextos.asp"   
	end sub
	
	sub deletar(id)
		dim sql
		sql = "delete from home_textos where idtexto = " & id
		call exec(sql)
		response.redirect "frmcadtextos.asp"   
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
			call getTextoById(id)
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
       window.onload = function () {
           CKEDITOR.replace('editor1');
       };
    </script> 
<script>
	function validaForm() {
		var form = document.frmcadtextos;
		if (form.cbarea.value == "") {
			alert("Por favor escolha uma Area");
			return false;
		}
		return true
	}
</script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<link href="../fckeditor/_samples/sample.css" rel="stylesheet" type="text/css" />
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="" name="frmcadtextos" method="POST" onSubmit="return validaForm()">
			<input type="hidden" name="id" value="<%=id%>" />
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<div id="painelcontrole">
						<table cellspacing="3" cellpadding="2" width="100%" border="0" id="tablelisttransportadoras">
							<tr>
								<td colspan="3" id="explaintitle" align="center">Cadastro de Textos do Site</td>
							</tr>
							<tr>
								<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td colspan="2" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">Texto:</td>
							</tr>	
							<tr>
                                <td colspan="3"><textarea id="editor1" name="editor1" value="<%=texto%>"><%=getText()%></textarea></td>
								<!--<td colspan="2"><%=getText()%></td>-->
							</tr>
							<tr>
								<td align="left" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">
									Area:
									<select name="cbarea" class="select">
										<option value="">[Selecione]</option>
										<option value="Home" <%if area = "home" then%>selected<%end if%>>Home</option>
										<!--option value="Cadastro" <%if area = "cadastro" then%>selected<%end if%>>Cadastro</option-->
										<!--option value="solicitacao">Solicita&ccedil;&atilde;o</option-->
										<option value="Regulamento" <%if area = "regulamento" then%>selected<%end if%>>Regulamento</option>
										<option value="PoliticaAmb" <%if area = "PoliticaAmb" then%>selected<%end if%>>Politica Ambiental</option>
										<!--option value="Corporativo" <%if area = "corporativo" then%>selected<%end if%>>Corporativo</option-->
										<option value="Preserva" <%if area = "preserva" then%>selected<%end if%>>Preserva</option>
										<option value="Coleta" <%if area = "coleta" then%>selected<%end if%>>Coleta</option>
										<option value="Residuos" <%if area = "residuos" then%>selected<%end if%>>Residuos</option>
									</select>
								</td>
							</tr>
							<tr>
								<%if len(trim(request.querystring("id"))) > 0 then%>
									<td align="right">
										<input type="submit" name="action" class="btnform" value="Editar" />
										<input type="submit" name="action" class="btnform" value="Deletar" />
									</td>
								<%else%>	
									<td align="right"><input type="submit" name="action" class="btnform" value="Salvar" onClick="validaForm()" /></td>
								<%end if%>
							</tr>
							<tr>
								<td align="center" width="100%">
									<table width="100%" cellspacing="1" cellpadding="1" id="tableRelSolPendente" style="border:1px solid #000000">
										<tr>
											<th width="2%"><img src="img/check.gif" /></th>
											<th>Texto</th>
											<th>Area</th>
										</tr>
										<%=getTextos()%>
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
