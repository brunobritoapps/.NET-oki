<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	function getListFiles()
		dim files
		dim fso
		dim folder
		dim ofolder
		dim nome_file
		dim i
		dim style
		dim nome
		dim datecreatedfile

		i = 0
		folder = request.servervariables("APPL_PHYSICAL_PATH")&"adm/transportadora/arquivos_gerados/"
		set fso = server.createobject("scripting.filesystemobject")
		set ofolder = fso.getfolder(folder)
		set files = ofolder.files
		for each nome_file in files
			if i mod 2 = 0 then
				style = "class=""classColorRelPar"""
			else
				style = "class=""classColorRelImpar"""
			end if	
			i = i + 1
			datecreatedfile = nome_file.dateCreated
			nome = split(nome_file, "\")
			
				getListFiles = getListFiles & "<tr>"
				if left(request.servervariables("LOCAL_ADDR"),3) = "127" then
					getListFiles = getListFiles & "<td "&style&" width=""2%""><input type=""checkbox"" name=""files"" id=""transp"&i&""" value="""&nome(7)&""" /></td><td "&style&"> <img src=""img/buscar.gif"" align=""absmiddle"" class=""imgexpandeinfo"" alt=""Download do Arquivo ["&nome(7)&"]"" onclick=""javascript:window.location='http://localhost:81/sgrs/adm/transportadora/arquivos_gerados/"&nome(7)&"';"" /> - <b>"&nome(6)&"</b> / "&nome(7)&"</td><td "&style&">"&datecreatedfile&"</td>"
				else
					getListFiles = getListFiles & "<td "&style&" width=""2%""><input type=""checkbox"" name=""files"" id=""transp"&i&""" value="""&nome(7)&""" /></td><td "&style&"> <img src=""img/buscar.gif"" align=""absmiddle"" class=""imgexpandeinfo"" alt=""Download do Arquivo ["&nome(7)&"]"" onclick=""javascript:window.location='http://www.sustentabilidadeoki.com.br/adm/transportadora/arquivos_gerados/"&nome(7)&"';"" /> - <b>"&nome(6)&"</b> / "&nome(7)&"</td><td "&style&">"&DateRight(formatdatetime(datecreatedfile,2))&"</td>"
				end if	
				getListFiles = getListFiles & "</tr>"
		next
		
		getListFiles = getListFiles & "<input type=""hidden"" name=""intsol"" value="""&i&""" />"
		
		set fso = nothing
		set ofolder = nothing
		set folder = nothing
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
		DateRight = Mes & "/" & Dia & "/" & Ano
	End Function
	
	sub deleteFiles()
		dim files
		dim fso
		dim file_name
		dim cont_files
		dim folder
		set fso = server.createobject("scripting.filesystemobject")
		files = split(request.form("files"), ",")
		for cont_files=0 to ubound(files)
			folder = server.MapPath("transportadora/arquivos_gerados/") & "\"
'			response.write folder & files(cont_files) & "<br />"
'			response.end
			folder = folder & replace(trim(files(cont_files))," ","")
'			response.write folder
			set file_name = fso.getfile(folder)
			file_name.delete
		next
		
		set file_name = nothing
		set fso = nothing
	end sub
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		call deleteFiles()
	end if
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
	function checkAll() {
		if (parseInt(document.frmlistafiletranspexport.intsol.value) > -1) {
			if (document.frmlistafiletranspexport.checkall.checked) {
				for (var i=1; i <= parseInt(document.frmlistafiletranspexport.intsol.value); i++) {
					var id = "transp"+i;
					document.getElementById(id).checked = true;
				}
			} else {
				for (var i=1; i <= parseInt(document.frmlistafiletranspexport.intsol.value); i++) {
					var id = "transp"+i;
					document.getElementById(id).checked = false;
				}
			}
		}
	}
</script>
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="frmlistafiletranspexport.asp" name="frmlistafiletranspexport" method="POST">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableprodcad">
						<tr>
							<td colspan="12" id="explaintitle" align="center">Listagem de Arquivo Eletrônico [Exportado] da Transportadora</td>
						</tr>
						<tr>
							<td colspan="12" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td align="right">Folder:</td>
							<td><b>Transportadora</b></td>
						</tr>
						<tr>
							<td colspan="12"><input type="checkbox" name="checkall" value="true" onClick="checkAll()" /> Selecionar Todos&nbsp;</td>
						</tr>
						<tr>
							<td colspan="12" align="center"><input name="apagar" type="submit" class="btnform" value="Apagar Arquivos" /></td>
						</tr>
						<tr>
							<td colspan="12">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro">
									<tr>
										<th><img src="img/check.gif" /></th>
										<th>Nome</th>
										<th>Data Criação</th>
									</tr>
									
									<%=getListFiles()%>
								</table>
							</td>
						</tr>
					</table>
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
