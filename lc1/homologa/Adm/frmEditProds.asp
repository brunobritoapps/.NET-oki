<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
Dim lIds, id
Dim sql, arr, intarr, i, v
Dim Msg

id = request("id")
Msg = request("Msg")

'Response.Write(Session.SessionID)

sub requests()
	lIds = request("lIds")
	qtd_check	= request.form("quantidade")
	qtd_check = split(qtd_check, ",")
	pont_check = request.form("pontuacao")
	pont_check = split(pont_check, ",")
	pont_tgt_check = request.form("pontuacaotarget")
	pont_tgt_check = split(pont_tgt_check, ",")
end sub
'======================================================================================================================================================================
sub addBonusTemp()
	dim sql
	dim quantidade
	dim pontuacao
	dim pontuacao_target

	v = split(lIds, ",")
	for i=0 to ubound(v)
		if len(trim(v(i))) > 0 then
			quantidade = request("quantidade_"&v(i))
			pontuacao = request("pontuacao_"&v(i))
			pontuacao_target = request("pontuacaotarget_"&v(i))
    
            sql = "select * from [Cadastro_bonus_has_produtos] where idoki_prod='" & trim(v(i)) & "' and SessionId=" & trim(Session.SessionID)
			'sql = "select * from [Cadastro_bonus_has_produtos_temp] where idoki_prod='" & trim(v(i)) & "' and SessionId=" & trim(Session.SessionID) 'peterson aquino 17-5-2014 id:6
'response.write sql & "<br>yyy"
'response.end
			call search(sql, arr, intarr)

			if clng(intarr) = -1 then
				'sql = "INSERT INTO [marketingoki2].[dbo].[Cadastro_bonus_has_produtos_temp] " & _ 'peterson aquino 17-5-2014 id:6
                sql = "INSERT INTO [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] " & _
							   "([idoki_prod] " & _
							   ",[qtd] " & _
							   ",[pontuacao] " & _
							   ",[pontuacao_target] " & _
							   ",[SessionId]) " & _
						 "VALUES " & _
							   "('"&trim(v(i))&"' " & _
							   ","&cint(quantidade)&" " & _
							   ","&ReplaceVirgula(pontuacao)&" " & _
							   ","&ReplaceVirgula(pontuacao_target)&" " & _
							   ",'"&trim(Session.SessionID)&"')"
				Response.Write sql & "<hr>"
				'Response.End
				call exec(sql)
			else
				'Response.Write "Passou aqui"
				Response.Redirect "frmEditProds.asp?Msg=Esse produto já foi adicionado."
			end if
		end if
	next	
end sub

sub addBonus()
	dim sql
	dim quantidade
	dim pontuacao
	dim pontuacao_target

	v = split(lIds, ",")
	for i=0 to ubound(v)
		if len(trim(v(i))) > 0 then
			quantidade = request("quantidade_"&v(i))
			pontuacao = request("pontuacao_"&v(i))
			pontuacao_target = request("pontuacaotarget_"&v(i))

			sql = "select * from [Cadastro_bonus_has_produtos] where idoki_prod='" & trim(v(i)) & "' and cad_cod_bonus='" & request("id") & "'"
'response.write sql & "<br>xxx"
'response.end
			call search(sql, arr, intarr)

			if clng(intarr) <= -1 then
				sql = "INSERT INTO [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] " & _
							   "([idoki_prod] " & _
							   ",[qtd] " & _
							   ",[pontuacao] " & _
							   ",[pontuacao_target] " & _
							   ",[cad_cod_bonus]) " & _
						 "VALUES " & _
							   "('"&trim(v(i))&"' " & _
							   ","&cint(quantidade)&" " & _
							   ","&ReplaceVirgula(pontuacao)&" " & _
							   ","&ReplaceVirgula(pontuacao_target)&" " & _
							   ",'"&trim(request("id"))&"')"

				call exec(sql)
			else
				sql = "UPDATE [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] SET " & _							   
							   " [qtd] = "&cint(quantidade)& _
							   ", [pontuacao] = "&ReplaceVirgula(pontuacao) & _
							   ", [pontuacao_target] = "&ReplaceVirgula(pontuacao_target) & _
							   ", [cad_cod_bonus] = '"&trim(request("id"))&"'" & _
						 " WHERE [idoki_prod] = '"&trim(v(i))&"'"

				'response.write sql
				'response.end
				call exec(sql)
				'Response.Redirect "frmEditProds.asp?Msg=Esse produto já foi adicionado."
			end if
		end if
	next	
end sub

function ReplaceVirgula(sValor)
	ReplaceVirgula = replace(sValor, ",",".")
end function

sub submit()
	if request.servervariables("HTTP_METHOD") = "POST" then
		call requests()
		
		if len(trim(request("id"))) <> 0 then
			call addBonus()
		else
			call addBonusTemp()
		end if
		
		Response.Redirect "frmEditProds.asp?Msg=Adicionado com sucesso."
	end if
end sub

call submit()

function PreencheValor(IdOki, lTipo)
	Dim sTipo
	Dim sql1, arr1, intarr1
	
	if len(trim(request("id"))) <> 0 then
		select case lTipo
			case 1 'quantidade
				sTipo = "qtd"
			case 2 'pontuacao
				sTipo = "pontuacao"
			case 3 'pontuacaotarget
				sTipo = "pontuacao_target"
		end select
		sql1 = "select " & sTipo & " from [Cadastro_bonus_has_produtos] where cad_cod_bonus='" & request("id") & "' and IdOki_Prod='" & IdOki & "'"

		'Response.Write sql1 & "<hr>"
		'Response.End

		call search(sql1, arr1, intarr1)

		'Response.Write clng(intarr1) & "<hr>"
		'Response.End

		if clng(intarr1) <> -1 then
			PreencheValor = arr1(0,intarr1)
		else
			PreencheValor = ""
		end if
	else
		select case lTipo
			case 1 'quantidade
				sTipo = "qtd"
			case 2 'pontuacao
				sTipo = "pontuacao"
			case 3 'pontuacaotarget
				sTipo = "pontuacao_target"
		end select
         sql1 = "select " & sTipo & " from [Cadastro_bonus_has_produtos] where IdOki_Prod='" & IdOki & "'"
		'sql1 = "select " & sTipo & " from [Cadastro_bonus_has_produtos_temp] where IdOki_Prod='" & IdOki & "'" 'peterson aquino 17-5-2014 id:6

		'Response.Write sql1 & "<hr>"
		'Response.End

		call search(sql1, arr1, intarr1)

		'Response.Write clng(intarr1) & "<hr>"
		'Response.End

		if clng(intarr1) <> -1 then
			PreencheValor = arr1(0,intarr1)
		else
			PreencheValor = ""
		end if
	end if
end function
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<script>
function validaForm() {
	var form = document.frmEditProds;
	var contObject = 0;
	var msg = "";
	var erro = false;
	var parametros=document.getElementById("lIds").value;
	var quebra=parametros.split(",");
	
	for (var i=0;i < document.getElementById("totalprodutos").value;i++) {
		//alert(eval('form.pontuacaotarget_'+quebra[i]+'.value'));
		if ((eval('form.quantidade_'+quebra[i]+'.value') == "" || eval('form.quantidade_'+quebra[i]+'.value') == 0) || (eval('form.pontuacao_'+quebra[i]+'.value') == "" || eval('form.pontuacao_'+quebra[i]+'.value') == 1) || (eval('form.pontuacaotarget_'+quebra[i]+'.value') == "" || eval('form.pontuacaotarget_'+quebra[i]+'.value') == 1)) {
			msg = msg + "Preencha os campos: quantidade, pontuacao e pontuacao target do produto corretamente\n";
			erro = true;
		}
		if (erro) {
			alert(msg);
			return;
		}
		/*
		if ((form.quantidade[i].value == "" || form.quantidade[i].value == 0) || (form.pontuacao[i].value == "" || form.pontuacao[i].value == 1) || (form.pontuacaotarget[i].value == "" || form.pontuacaotarget[i].value == 1)) {
			msg = msg + "Preencha os campos: quantidade, pontuacao e pontuacao target do produto corretamente\n";
			erro = true;
		}
		if (erro) {
			alert(msg);
			return;
		}
		*/
	}
	form.submit();
}
</script>
</head>
<body>
<%
if len(trim(Msg)) > 0 then
	Response.Write Msg
else
	if len(request("lIds")) > 0 then
		lIds = mid(request("lIds"),1,len(request("lIds"))-1)
		lIds = "'" & replace(lIds, ",", "','") & "'"
%>
	<form action="#" name="frmEditProds" method="POST">
	<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableGetClientesCadastro" style="border:1px solid #333333;">
		<tr style="background-color:#FF0000;">
			<td>ID</td>
			<td>Descrição</td>
			<td>quantidade</td>
			<td>pontuação</td>
			<td>pontuação target</td>
		</tr>
		<tr>
			<%		
			sql = "select IDOki, descricao from produtos where IDOki in (" & lIds & ")"
			
			call search(sql, arr, intarr)
			
			if clng(intarr) <> -1 then
				for i=0 to intarr
					if i mod 2 = 0 then
						style = " class=""classColorRelPar"""
					else
						style = " class=""classColorRelImpar"""
					end if
					Response.Write "<tr>" & vbcrlf
					Response.Write "<td"&style&" width=""120"">"&trim(arr(0,i))&"</td>" & vbcrlf
					Response.Write "<td"&style&">"&arr(1,i)&"</td>" & vbcrlf
					Response.Write "<td"&style&" width=""10""><input type=""text"" id=""quantidade_"&trim(arr(0,i))&""" name=""quantidade_"&trim(arr(0,i))&""" class=""textreadonly"" value="""&PreencheValor(trim(arr(0,i)),1)&""" size=""10"" /></td>" & vbcrlf
					Response.Write "<td"&style&" width=""10""><input type=""text"" id=""pontuacao_"&trim(arr(0,i))&""" name=""pontuacao_"&trim(arr(0,i))&""" class=""textreadonly"" value="""&PreencheValor(trim(arr(0,i)),2)&""" size=""10"" /></td>" & vbcrlf
					Response.Write "<td"&style&" width=""10""><input type=""text"" id=""pontuacaotarget_"&trim(arr(0,i))&""" name=""pontuacaotarget_"&trim(arr(0,i))&""" class=""textreadonly"" value="""&PreencheValor(trim(arr(0,i)),3)&""" size=""10"" /></td>" & vbcrlf
					Response.Write "</tr>" & vbcrlf
				next
				Response.Write "<input type=""hidden"" name=""totalprodutos"" id=""totalprodutos"" value="""&i&""" />"
			end if
			%>
			<!--
			<td>1</td>
			<td>2</td>
			<td width="10"><input type="text" id="quantidade" name="quantidade" class="textreadonly" value="" disabled="disabled" size="10" /></td>
			<td width="10"><input type="text" id="pontuacao" name="pontuacao" class="textreadonly" value="" disabled="disabled" size="10" /></td>
			<td width="10"><input type="text" id="pontuacaotarget" name="pontuacaotarget" class="textreadonly" value="" disabled="disabled" size="10" /></td>
			-->
		</tr>
		<tr>
			<td colspan=5 align=right>
				<INPUT type="hidden" value="<%=replace(lIds, "'", "")%>" id=lIds name=lIds>
				<INPUT type="button" value="Adicionar" onClick="validaForm();">
			</td>
		</tr>
	</table>
	</form>
	<%end if%>
<%end if%>
</body></html>
