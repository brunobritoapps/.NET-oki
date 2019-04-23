<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	dim cod_bonus
	dim desc_bonus
	dim validade
	dim moeda
	dim aplicacao
	dim data_inicio_cont
	dim prod_check
	dim qtd_check
	dim pont_check
	dim pont_tgt_check
	redim produtos(0,0)

	sub getBonus(id)
		dim sql, arr, intarr, i
		sql = "SELECT [cod_bonus] " & _
					  ",[descricao] " & _
					  ",[validade] " & _
					  ",[moeda] " & _
					  ",[aplicacao] " & _
					  ",[data_inicio_contabilizacao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_Bonus] where [cod_bonus] = '"&id&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				cod_bonus = arr(0,i)
				desc_bonus = arr(1,i)
				validade = arr(2,i)
				moeda = arr(3,i)
				aplicacao = arr(4,i)
				data_inicio_cont = arr(5,i)
			next
			call getProdutosByBonus(id)
		else
			response.write "<script>alert('"&id&" n�o encontrado.')</script>"
		end if
	end sub

	function getProdutosByBonus(id)
		dim sql, arr, intarr, i

		sql = "SELECT a.[idoki_prod] " & _
				",a.[qtd] " & _
				",a.[pontuacao] " & _
				",a.[pontuacao_target] " & _
				",a.[cad_cod_bonus] " & _
				",b.[descricao] " & _
				"FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] as a " & _
				"left join [marketingoki2].[dbo].[Produtos] as b " & _
				"on a.[idoki_prod] = b.[IDOki] " & _
				"where a.[cad_cod_bonus] = '"&id&"'"

		call search(sql, arr, intarr)
		if intarr > -1 then
			redim produtos(5,intarr)
			for i=0 to intarr
				produtos(0,i) = arr(0,i)
				produtos(1,i) = arr(5,i)
				produtos(2,i) = arr(1,i)
				produtos(3,i) = arr(2,i)
				produtos(4,i) = arr(3,i)
			next
		end if
	end function

	Function IncrementNumber(StringNumber)
		Dim DiffNumber
		Dim NewNumber
		Dim i
		Dim sSql

		StringNumber = StringNumber + 1
		DiffNumber = 5 - Len(StringNumber)
		For i=1 To DiffNumber
			NewNumber = NewNumber & "0"
		Next
		NewNumber = NewNumber & StringNumber
		IncrementNumber = NewNumber
	End Function

	Function getSequencialBonus()
		Dim numSequencial, dataNumSequencial
		Dim sSql, arrNumSeqIDCliente, intNumSeqIDCliente, i

		sSql = "select count([cod_bonus]) from cadastro_bonus"
		Call search(sSql, arrNumSeqIDCliente, intNumSeqIDCliente)
		If intNumSeqIDCliente > -1 Then
			if arrNumSeqIDCliente(0,i) <> 0 and arrNumSeqIDCliente(0,i) <> "" then
				For i=0 To intNumSeqIDCliente
					numSequencial = IncrementNumber(arrNumSeqIDCliente(0,i))
				Next
			else
				numSequencial = "00001"
			end if
		Else
			numSequencial = "00001"
		End If
		getSequencialBonus = numSequencial
	End Function

	function comparaProduto(id)
		dim i
		dim booleano
		i=0
		booleano = false
		while (i <= ubound(produtos,2) and booleano = false)
			if trim(produtos(0,i)) = trim(id) then
				comparaProduto = "checked=""checked"""
				booleano = true
			else
				comparaProduto = ""
			end if
			i = i + 1
		wend
	end function

	function comparaProduto2(id)
		dim i
		dim booleano
		i=0
		booleano = false
		while (i <= ubound(produtos,2) and booleano = false)
			if trim(produtos(0,i)) = trim(id) then
				comparaProduto2 = i
				booleano = true
			else
				comparaProduto2 = -1
			end if
			i = i + 1
		wend
	end function

	function getListProdutos()
		dim sql, arr, intarr, i
		dim arrgrupos, intgrupos, j, n
		dim ret
		dim style
		dim id_prod_equal
		dim quantidade
		dim cont
		dim contadorRegistros
		dim soma
		ret = ""
		style = "class=""classColorRelPar"""
		soma = 0
		sql = "SELECT [idGrupo_produtos] " & _
		  ",[descricao] " & _
		  ",[idokigrupo] " & _
	  "FROM [marketingoki2].[dbo].[Grupo_produtos]"
		call search(sql, arrgrupos, intgrupos)
		if intgrupos > -1 then
			for n=0 to intgrupos

				sql = "select idoki, descricao from produtos where gera_bonus = 1 and grupo_produtos_idgrupo_produtos = " & arrgrupos(0,n)
			
				call search(sql, arr, intarr)
				if clng(intarr) <> -1 then
					soma = soma + intarr
				end if
				if intarr > -1 then
					ret = ret & "<div align=""left"" style=""background-color:#FF0000;padding:5px 5px 5px 5px;border-bottom:1px solid #CCCCCC;""><img src=""img/+.gif"" id="""&trim(arrgrupos(1,n))&""" class=""imgexpandeinfo"" onclick=""atualizaDisplayGrupo('"&trim(arrgrupos(0,n))&"','"&trim(arrgrupos(1,n))&"')"" /> <b style=""color:#FFFFFF;"">"&trim(arrgrupos(1,n))&"</b> <input type=""checkbox"" name=""checkGroup"&n&""" onclick=""verificaGrupo('idProdByGroup"&n&"', 'checkGroup"&n&"', '"&trim(arrgrupos(0,n))&"', '"&trim(arrgrupos(1,n))&"')"" /></div>"
					ret = ret & "<div id="""&trim(arrgrupos(0,n))&""" style=""display:none;"">"
					ret = ret & "<table cellpadding=""1"" cellspacing=""1"" width=""100%"" align=""center"" id=""tableGetClientesCadastro"" style=""border:1px solid #333333;"">"
					ret = ret & "<tr>"
					ret = ret & "<th width=""15""><img src=""img/check.gif"" /></th>"
					ret = ret & "<th>ID</th>"
					ret = ret & "<th>Descri��o</th>"
					ret = ret & "<th>quantidade</th>"
					ret = ret & "<th>pontua��o</th>"
					ret = ret & "<th>pontua��o target</th>"
					ret = ret & "</tr>"
					for i=0 to intarr
						if i mod 2 = 0 then
							style = "class=""classColorRelPar"""
						else
							style = "class=""classColorRelImpar"""
						end if
						if request.querystring("id") <> "" then
							ret = ret & montaHtmlTabela(arr, i, trim(arrgrupos(0,n)))
						else
							ret = ret & "<tr>"
							ret = ret & "<input type=""hidden"" name=""idProdByGroup"&n&""" value="""&getIDProdByGroup(trim(arrgrupos(0,n)), i)&""" />"
							ret = ret & "<td width=""5"" "&style&"><input type=""checkbox"" name=""checkar"" id=""checkar_"&trim(arr(0,i))&""" value="""&trim(arr(0,i))&""" onClick=""checkProduto('checkar_"&trim(arr(0,i))&"','pontuacao_"&trim(arr(0,i))&"','quantidade_"&trim(arr(0,i))&"','pontuacaotarget_"&trim(arr(0,i))&"')"" /></td>"
							ret = ret & "<td width=""120"" "&style&">"&trim(arr(0,i))&"</td>"
							ret = ret & "<td "&style&">"&arr(1,i)&"</td>"
							ret = ret & "<td width=""10"" "&style&"><input type=""text"" id=""quantidade_"&trim(arr(0,i))&""" name=""quantidade"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
							ret = ret & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacao_"&trim(arr(0,i))&""" name=""pontuacao"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
							ret = ret & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacaotarget_"&trim(arr(0,i))&""" name=""pontuacaotarget"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
							ret = ret & "</tr>"
						end if
					next
					ret = ret & "</table>"
					ret = ret & "</div>"
				end if
			next
			ret = ret & "<input type=""hidden"" name=""totalprodutos"" id=""totalprodutos"" value="""&soma&""" />"
		end if
		getListProdutos = ret
	end function

	function getListGruposNew()
		dim sql, arr, intarr, i
		dim arrgrupos, intgrupos, j, n
		dim ret
		dim id_prod_equal
		dim quantidade
		dim cont
		dim contadorRegistros
		ret = ""

		sql = "SELECT [idGrupo_produtos] " & _
		  ",[descricao] " & _
	  "FROM [marketingoki2].[dbo].[Grupo_produtos]"

		call search(sql, arrgrupos, intgrupos)

		if intgrupos > -1 then
			for n=0 to intgrupos
				ret = ret & "<option value="""&arrgrupos(0,n)&""">"&arrgrupos(1,n)&"</option>"
			next
			ret = ret & ""
		end if
		getListGruposNew = ret
	end function
	
	function getListProdutosNew()
		dim sql, arr, intarr, i
		dim arrgrupos, intgrupos, j, n
		dim ret
		dim id_prod_equal
		dim quantidade
		dim cont
		dim contadorRegistros
		ret = ""

		Response.Write request("IdGrupo") & "<hr>"
		if len(trim(request("IdGrupo"))) > 0 then
			sql = "select idoki, descricao from produtos where gera_bonus = 1 and grupo_produtos_idgrupo_produtos in (" & request("IdGrupo") & ")"
			Response.Write sql & "<hr>"

			call search(sql, arrgrupos, intgrupos)

			if intgrupos > -1 then
				for n=0 to intgrupos
					ret = ret & "<option value="""&arrgrupos(0,n)&""">"&arrgrupos(1,n)&"</option>"
				next
				ret = ret & ""
			end if
			getListProdutosNew = ret
		end if
	end function

	function getIDProdByGroup(id, j)
		if j=0 then
			dim sql, arr, intarr, i
			dim retorno
			sql = "select idoki, descricao from produtos where gera_bonus = 1 and grupo_produtos_idgrupo_produtos = "&id
			call search(sql, arr, intarr)
			if intarr > -1 then
				for i=0 to intarr
					if i=0 then
						retorno = trim(arr(0,i))
					else
						retorno = retorno &","&trim(arr(0,i))
					end if
				next
			else
				retorno = ""
			end if
			getIDProdByGroup = retorno
		end if
	end function

	function montaHtmlTabela(arr, i, grupo)
		dim j
		dim html, style
		dim bExist, indice
		html = ""
		indice = -1

		if i mod 2 = 0 then
			style = "class=""classColorRelPar"""
		else
			style = "class=""classColorRelImpar"""
		end if
		html = html & "<tr>"
		html = html & "<td width=""5"" "&style&"><input type=""checkbox"" "&comparaProduto(arr(0,i))&" name=""checkar"" id=""checkar_"&trim(arr(0,i))&""" value="""&trim(arr(0,i))&""" onClick=""checkProduto('checkar_"&trim(arr(0,i))&"','pontuacao_"&trim(arr(0,i))&"','quantidade_"&trim(arr(0,i))&"','pontuacaotarget_"&trim(arr(0,i))&"')"" /></td>"
		html = html & "<td width=""120"" "&style&">"&trim(arr(0,i))&"</td>"
		html = html & "<td "&style&">"&trim(arr(1,i))&"</td>"
		indice = comparaProduto2(arr(0,i))
		if not indice <> -1 then
			html = html & "<input type=""hidden"" name=""idProdByGroup"&i&""" value="""&getIDProdByGroup(grupo, i)&""" />"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""quantidade_"&trim(arr(0,i))&""" name=""quantidade"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacao_"&trim(arr(0,i))&""" name=""pontuacao"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacaotarget_"&trim(arr(0,i))&""" name=""pontuacaotarget"" class=""textreadonly"" value="""" disabled=""disabled"" size=""10"" /></td>"
		else
			html = html & "<input type=""hidden"" name=""idProdByGroup"&i&""" value="""&getIDProdByGroup(grupo, i)&""" />"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""quantidade_"&trim(arr(0,i))&""" name=""quantidade"" class=""text"" value="""&trim(produtos(2,indice))&""" size=""10"" /></td>"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacao_"&trim(arr(0,i))&""" name=""pontuacao"" class=""text"" value="""&trim(produtos(3,indice))&""" size=""10"" /></td>"
			html = html & "<td width=""10"" "&style&"><input type=""text"" id=""pontuacaotarget_"&trim(arr(0,i))&""" name=""pontuacaotarget"" class=""text"" value="""&trim(produtos(4,indice))&""" size=""10"" /></td>"
		end if
		html = html & "</tr>"
		montaHtmlTabela = html
	end function

	sub requests()
		cod_bonus					= request.form("cod_bonus")
		desc_bonus					= request.form("textdesc")
		validade					= request.form("validade")
		moeda						= request.form("cbMoeda")
		aplicacao					= request.form("cbAplicacao")
		data_inicio_cont			= request.form("data_inicio")
		'prod_check					= request.form("checkar")
		'prod_check = split(prod_check, ",")
		'qtd_check					= request.form("quantidade")
		'qtd_check = split(qtd_check, ",")
		'pont_check					= request.form("pontuacao")
		'pont_check = split(pont_check, ",")
		'pont_tgt_check				= request.form("pontuacaotarget")
		'pont_tgt_check = split(pont_tgt_check, ",")
	end sub
'======================================================================================================================================================================
	sub addBonus()
		dim sql, arr, intarr
		dim quantidade
		dim pontuacao
		dim pontuacao_target
		dim exist
		exist = verifyBonus(cod_bonus)
'		response.write exist
'		response.end
		if not exist then
			cod_bonus = getSequencialBonus()
			sql = "INSERT INTO [marketingoki2].[dbo].[Cadastro_Bonus] " & _
						   "([cod_bonus] " & _
						   ",[descricao] " & _
						   ",[validade] " & _
						   ",[moeda] " & _
						   ",[aplicacao] " & _
						   ",[data_inicio_contabilizacao]) " & _
					 "VALUES " & _
						   "('"&cod_bonus&"' " & _
						   ",'"&desc_bonus&"' " & _
						   ","&validade&" " & _
						   ",'"&moeda&"' " & _
						   ",'"&aplicacao&"' " & _
						   ",CONVERT(DATETIME, '"&FormatDate(data_inicio_cont)&"'))"
		else
			sql = "UPDATE [marketingoki2].[dbo].[Cadastro_Bonus] " & _
				   "SET [descricao] = '"&desc_bonus&"' " & _
					  ",[validade] = "&validade&" " & _
					  ",[moeda] = '"&moeda&"' " & _
					  ",[aplicacao] = '"&aplicacao&"' " & _
					  ",[data_inicio_contabilizacao] = CONVERT(DATETIME, '"&FormatDate(data_inicio_cont)&"') " & _
				 "WHERE [cod_bonus] = '"&cod_bonus&"'"
		end if
'		response.write sql
'		response.end
		call exec(sql)
		dim i
		
		'aqui q vou mudar a forma de inclus�o, tem q buscar da tabela temp
		sql = "select * from Cadastro_bonus_has_produtos_temp where SessionId = " & Session.SessionID

		'Response.Write sql & "<hr>"
		'Response.End

		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to ubound(arr)
				call deletaProdutosByBonus(cod_bonus, arr(0, intarr))
			next			
			
			sql = "INSERT INTO [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] " & _
						   "([idoki_prod] " & _
						   ",[qtd] " & _
						   ",[pontuacao] " & _
						   ",[pontuacao_target] " & _
						   ",[cad_cod_bonus]) " & _
					 	   "select [idoki_prod] ,[qtd] ,[pontuacao] ,[pontuacao_target], '"&trim(cod_bonus)&"' as cad_cod_bonus from Cadastro_bonus_has_produtos_temp where sessionid="&Session.SessionID
			'Response.Write sql & "<hr>"
			'Response.End
			call exec(sql)
			call exec("delete from Cadastro_bonus_has_produtos_temp where sessionid="&Session.SessionID)
		else
		end if

	end sub
'======================================================================================================================================================================
	sub deletaProdutosByBonus(id, idprod)
		dim sqlDel, arrDel, intarrDel
		sqlDel = "select * from cadastro_bonus_has_produtos where [cad_cod_bonus] = '"&id&"' and idoki_prod = '"&idprod&"'"
		call search(sqlDel, arrDel, intarrDel)
		if intarrDel <> -1 then
			call exec("delete from cadastro_bonus_has_produtos where [cad_cod_bonus] = '"&id&"' and idoki_prod = '"&idprod&"'")
			'response.write "deleta produtos: " & sql & "<br />"			
		end if
	end sub
'======================================================================================================================================================================
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

	function verifyBonus(cod)
		dim sql, arr, intarr
		sql = "select * from cadastro_bonus where cod_bonus = '"&cod&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			verifyBonus = true
		else
			verifyBonus = false
		end if
	end function

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
		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function

	sub submit()
		if request.servervariables("HTTP_METHOD") = "POST" then
			call requests()
			call addBonus()
			cod_bonus					= ""
			desc_bonus					= ""
			validade					= ""
			moeda						= -1
			aplicacao					= -1
			data_inicio_cont			= ""
		else
			if request.querystring("id") <> "" then
				call getBonus(request.querystring("id"))
			end if
		end if
	end sub

	call submit()
%>
<html>
<head>
<script>
	function checkProduto(name, pontuacao, quantidade, pontuacaotarget) {
		var form = document.frmcadbonus;
		if (document.getElementById(name).checked) {
			document.getElementById(pontuacao).disabled 		= '';
			document.getElementById(quantidade).disabled 		= '';
			document.getElementById(pontuacaotarget).disabled 	= '';
			document.getElementById(pontuacao).className 		= 'text';
			document.getElementById(quantidade).className 		= 'text';
			document.getElementById(pontuacaotarget).className = 'text';
		} else {
			document.getElementById(pontuacao).disabled 		= 'disabled';
			document.getElementById(quantidade).disabled 		= 'disabled';
			document.getElementById(pontuacaotarget).disabled 	= 'disabled';
			document.getElementById(pontuacao).className 		= 'textreadonly';
			document.getElementById(quantidade).className 		= 'textreadonly';
			document.getElementById(pontuacaotarget).className = 'textreadonly';
		}
	}

	function atualizaDisplayGrupo(id, img) {
		document.getElementById(id).style.display = (document.getElementById(id).style.display == 'block')?document.getElementById(id).style.display = 'none':document.getElementById(id).style.display = 'block';
		document.getElementById(img).src = (document.getElementById(img).src == 'http://200.225.91.166/sgrs//adm/img/+.gif')?document.getElementById(img).src = 'http://200.225.91.166/sgrs//adm/img/-.gif':document.getElementById(img).src = 'http://200.225.91.166/sgrs//adm/img/+.gif';
	}

	function date(campo) {
			var string = campo.value;
			var _char = "/";
			for (var i=0;i<string.length;i++) {
				if (i == 2 || i == 5) {
					continue;
				}
			}
			switch (string.length) {
				case 2:
					string += _char;
					break;
				case 5:
					string += _char;
					break;
			}
			campo.value = string;
	}

	function validaForm() {
		var form = document.frmcadbonus;
		var contObject = 0;
		var msg = "";
		var erro = false;
/*
		for (var i=0;i < document.getElementById("totalprodutos").value;i++) {
			if (!form.checkar[i].checked) {
				contObject++;
			} else {
				if ((form.quantidade[i].value == "" || form.quantidade[i].value == 0) || (form.pontuacao[i].value == "" || form.pontuacao[i].value == 1) || (form.pontuacaotarget[i].value == "" || form.pontuacaotarget[i].value == 1)) {
					msg = msg + "Preencha os campos: quantidade, pontuacao e pontuacao target do produto["+form.checkar[i].value+"] corretamente\n";
					erro = true;
				}
			}
			if (contObject == parseInt(document.getElementById("totalprodutos").value)) {
				alert("Por favor escolha um Produto para esse B�nus");
				return;
			}
			if (erro) {
				alert(msg);
				return;
			}
		}
*/		
		if (form.textdesc.value == "") {
			alert("Por favor preencha o campo Descri��o");
			return;
		}
		if (form.validade.value == "") {
			alert("Por favor preencha o campo Validade");
			return;
		}
		if (form.cbMoeda.value == -1) {
			alert("Por favor escolha um tipo de Moeda");
			return;
		}
		if (form.cbAplicacao.value == -1) {
			alert("Por favor escolha um tipo de Aplica��o");
			return;
		}
		if (form.data_inicio.value == "" || form.data_inicio.value == "___/___/___") {
			alert("Por favor preencha o in�cio da Contabiliza��o do B�nus");
			return;
		}
		form.submit();
	}

	function verificaGrupo(id, check, grupo, img) {
		var texto = document.getElementById(id).value;
		var idprod = "";
		var arraySplit = texto.split(',');
		atualizaDisplayGrupo(grupo, img);
		for (var i=0;i<arraySplit.length;i++) {
			idprod = 'checkar_'+arraySplit[i];
			document.getElementById(idprod).checked = (document.getElementById(idprod).checked)?false:true;
			checkProduto('checkar_'+arraySplit[i],'pontuacao_'+arraySplit[i],'quantidade_'+arraySplit[i],'pontuacaotarget_'+arraySplit[i]);
		}
	}

</script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
.controls {
	 border: solid 1px #CCCCCC;
	 border-top: solid 1px #CCCCCC;
	 border-top-width: 1px;
 	 border-bottom-width: 1px;
	 border-left-width: 1px;
	 border-right-width: 1px;
	 border-top-color: #CCCCCC;
	 border-bottom-color: #CCCCCC;
	 border-left-color: #CCCCCC;
	 border-right-color: #CCCCCC;
	 border-top-style: solid;
	 border-left-style: solid;
	 border-right-style: solid;"
 
}
</style>
<SCRIPT LANGUAGE="JavaScript" SRC="js/CalendarPopup.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var cal = new CalendarPopup();
</SCRIPT>
<script type="text/javascript">
//<![CDATA[

var xmlhttp = false;
if (window.XMLHttpRequest) {
   xmlhttp = new XMLHttpRequest(  );
   xmlhttp.overrideMimeType('text/xml');
} else if (window.ActiveXObject) {
   xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
}

function populateList() {

   var Grupo = frmcadbonus.Grupos.options.value;
   var Prod = frmcadbonus.txtSearchProd.value;
   //alert('ajax.asp?Sub=getprod&SearchProd='+Prod+'&IdGrupo=' + Grupo);
   var url = 'ajax.asp?Sub=getprod&SearchProd='+Prod+'&IdGrupo=' + Grupo;
   xmlhttp.open('GET', url, true);
   xmlhttp.onreadystatechange = getProds;
   xmlhttp.send(null);

}

function getProds() {
	 if(xmlhttp.readyState == 4) {
			document.getElementById('Prods').innerHTML = xmlhttp.responseText;
   } else {
      document.getElementById('Prods').innerHTML = 'Carregando...';
   }
}
//]]>

function MoveSelectedListItems(srcCombo)
{
	var numItems = 0;
	var srcLen = srcCombo.options.length;
	var lIds = '';
	
	for (var x=0; x<srcCombo.options.length; x++)
		if (srcCombo.options[x].selected) numItems++;
	
	for (var x=0; x<srcCombo.options.length; x++)
		if (srcCombo.options[x].selected==true) 
		{
			//alert(srcCombo.options[x].value);
			lIds = srcCombo.options[x].value + "," + lIds
		}
	
	//alert("frmEditProds.asp?Lids="+lIds+"&id=<%=request.querystring("id")%>");	//window.parent.EditProds.location.href="frmEditProds.asp?Lids="+lIds+"&id=<%'=request.querystring("id")%>";
	document.getElementById('EditProds').src = "frmEditProds.asp?Lids="+lIds+"&id=<%=request.querystring("id")%>";	
}
</script>



</head>
<body>
<div id="container">
  <!--#include file="inc/i_header.asp" -->
  <div id="conteudo">
    <table cellspacing="0" cellpadding="0" width="775" ID="Table1">
      <form action="frmcadbonus.asp" name="frmcadbonus" method="POST">
        <tr>
          <td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
          <td id="conteudo"><table cellspacing="1" cellpadding="1" width="100%" border=0 id="tableEditSolicitacaoColetaAdm">
              <tr>
                <td colspan="3" id="explaintitle" align="center">Cadastro de B�nus</td>
              </tr>
              <tr>
                <td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
              </tr>
              <tr>
                <td width="35%" align="right">C�d. B�nus:</td>
                <td><input type="text" name="cod_bonus" class="text" value="<%= cod_bonus %>" maxlength="10" readonly="readonly"/>
                  * </td>
              </tr>
              <tr>
                <td width="35%" align="right">Descri��o:</td>
                <td><textarea name="textdesc" cols="50" rows="10" class="textoHome"><%= desc_bonus %></textarea></td>
              </tr>
              <tr>
                <td width="35%" align="right">Validade:</td>
                <td><input type="text" name="validade" class="text" value="<%= validade %>" size="6" />
                  * </td>
              </tr>
              <tr>
                <td width="35%" align="right">Moeda:</td>
                <td><select name="cbMoeda" class="select">
                    <option value="-1">[Selecione]</option>
                    <option value="P" <%if moeda = "P" then%>selected<%end if%>>PONTUA��O</option>
                    <option value="R" <%if moeda = "R" then%>selected<%end if%>>REAL</option>
                    <option value="D" <%if moeda = "D" then%>selected<%end if%>>DOLAR</option>
                  </select>
                  * </td>
              </tr>
              <tr>
                <td width="35%" align="right">Aplica��o:</td>
                <td><select name="cbAplicacao" class="select">
                    <option value="-1">[Selecione]</option>
                    <option value="CLI" <%if aplicacao = "CLI" then%>selected<%end if%>>CLIENTE</option>
                    <option value="PONTO" <%if aplicacao = "PONTO" then%>selected<%end if%>>PONTO DE COLETA</option>
                  </select>
                  * </td>
              </tr>
              <tr>
                <td width="35%" align="right">Data in�cio Contabiliza��o:</td>
                <%if data_inicio_cont <> "" then%>
					<%if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then%>
						<td>
							<input type="text" name="data_inicio" class="text" readonly="true" maxlength="10" size="12" value="<%= DateRight(data_inicio_cont) %>" onKeyPress="date(this)" />
							<A HREF="#" onClick="cal.select(document.forms['frmcadbonus'].data_inicio,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img src="img/btn_calendario.gif" border="0"></A>
						  * </td>
					<%else%>
						<td>
							<input type="text" name="data_inicio" class="text" readonly="true" maxlength="10" size="12" value="<%= data_inicio_cont %>" onKeyPress="date(this)" />
							<A HREF="#" onClick="cal.select(document.forms['frmcadbonus'].data_inicio,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img src="img/btn_calendario.gif" border="0"></A>
						  * </td>
					<%end if%>
                <%else%>
                <td>
									<input type="text" name="data_inicio" class="text" readonly="true"  maxlength="10" size="12" value="" onKeyPress="date(this)" />
									<A HREF="#" onClick="cal.select(document.forms['frmcadbonus'].data_inicio,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img src="img/btn_calendario.gif" border="0"></A>
                  * </td>
                <%end if%>
              </tr>
              <tr>
                <td colspan="2" align="center">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2" align="center"><fieldset style="width:700px;padding:5px 5px 5px 5px;">
                  <legend style="color:#DE8989;font-weight:bold;">A��es</legend>
                  <div align="right">
                    <input type="button" class="btnform" name="listar" value="Listar B�nus" onClick="javascript:window.open('frmlistabonus.asp','','width=700,height=400,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" />
                    <input type="button" class="btnform" name="cad_submit" value="<%if len(trim(request.querystring("id"))) <> 0 then%>Editar B�nus<%else%>Cadastrar B�nus<%end if%>" onClick="validaForm()" />
                  </div>
                  </fieldset></td>
              </tr>
              <tr>
                <td colspan="2" align="center"><fieldset style="width:700px;padding:5px 5px 5px 5px;">
                  <legend style="color:#DE8989;font-weight:bold;">Produtos</legend>

				  </
                  <br />
                  <!--
                  <input type="text" name=t1 value="" size=20>
									<input type="button" name=b1 value="Find" onClick="if(this.t1.value!=null && this.t1.value!='')findString(this.t1.value);return false">
									-->
                    <%
					'=getListProdutos()
					%>
                    
                    <table width=100% border=0>
											<tr>
												<td>
													<!--
													<INPUT type="text" name=txtSearchGrupo>													<INPUT type="button" value="Procurar">													-->
												</td>												<td>&nbsp;</td>												<td>
													Procure pelo C�digo ou Descri��o: <br>
													
										<INPUT type="text" name=txtSearchProd>													
										<INPUT type="button" value="Procurar" onClick="populateList();">
												</td>												<td>&nbsp;</td>												<td>&nbsp;</td>
											</tr>											<tr>												<td>
												
										<select name="Grupos" id="Grupos" onChange="populateList();" class="controls" style="width: 180px; height: 210px; font-size: 13px;" size="8">												<%=getListGruposNew()%>
												</SELECT>
												</td>
												<td>&nbsp;</td>
												<td>
												<div id="Prods">
												<select class="controls" style="width: 330px; height: 210px; font-size: 13px;" size="8">
												</select>
												</div>
												</td>
												<td>&nbsp;</td>
												<td><INPUT type="button" value="Listar" onClick="MoveSelectedListItems(document.frmcadbonus.Prods);"></td>
											</tr>
										</table>
										<table width=100% border=1>
											<tr>
												<td>
													<iframe name="EditProds" id="EditProds" src="frmEditProds.asp" width="100%" scrolling="auto" frameborder="0"></iframe>
												</td>
											</tr>
										</table>
                    
                  </fieldset></td>
              </tr>
            </table></td>
          <td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
        </tr>
      </form>
    </table>
  </div>
  <!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%
'Call close()
%>
