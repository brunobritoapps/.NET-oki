<!--#include file="../_config/_config.asp" -->
<link href="../css/geral.css" rel="stylesheet" type="text/css"/>
<%
'|--------------------------------------------------------------------
'| Arquivo: frmCartaDoacaoNF.asp
'| Autor: Jadilson Muramatsu (jmuramatsu@hotmail.com)
'| Data Cria��o: 09/10/2007
'| Data Modifica��o : 09/10/2007
'| Descri��o: Gera��o da carta de doacao do cliente (ASP)
'|--------------------------------------------------------------------
%>
<%Call open()%>
<%
	'============================================================================================
	Dim sTxt, sAcao
	dim lIdSolicitacaoColeta, lTipoPessoa
	
	Dim lNrDoc, sRemetente, sCNPJ, sEndereco, sCidade, sCEP, sUF, lQtd, sNrSolCol, lTipoColeta
	Dim lIdPontoColeta
	Dim lAdm
	
	sAcao = request("Acao")
	lIdSolicitacaoColeta = request("IdSolicitacaoColeta")
	lTipoPessoa = request("TipoPessoa")
	lAdm = request("Adm")
	
	'============================================================================================

	if sAcao = "1" then
		call geraCartaDoacao()
	else
		lTipoColeta	= request("TipoColeta")
	end if
	
	'============================================================================================
	'| Sub que gera as Categorias para o cliente Selecionar
	'============================================================================================
	Sub geraCartaDoacao()
		Dim sSql, arrCarta, intCarta, i

		'sSql = "SELECT S.Solicitacao_coleta_idSolicitacao_Coleta, C.Razao_Social, C.CNPJ, "&_
		'			 "CEP.Logradouro, C.Compl_Endereco, C.Numero_Endereco, "&_
		'			 "CEP.Municipio, CEP.cep, CEP.Estado, COL.qtd_cartuchos, COL.numero_solicitacao_coleta, C.typeColect, S.Pontos_Coleta_idPontos_Coleta from clientes C "&_
		'			 "INNER JOIN Solicitacao_coleta_has_Clientes S ON C.IdClientes = S.Clientes_IdClientes "&_
		'			 "INNER JOIN Solicitacao_coleta COL ON S.Solicitacao_coleta_idSolicitacao_coleta = COL.idSolicitacao_coleta "&_
		'			 "INNER JOIN cep_consulta_has_Clientes CEP ON C.IdClientes = CEP.Clientes_IdClientes "&_
		'			 "WHERE COL.idSolicitacao_coleta = " & lIdSolicitacaoColeta
					 
		sSql = "SELECT S.Solicitacao_coleta_idSolicitacao_Coleta, PC.Razao_Social, PC.CNPJ, "&_
				"PC.Logradouro, PC.Complemento_Endereco, PC.Numero_Endereco, "&_
				"PC.Municipio, PC.cep, PC.Estado, COL.qtd_cartuchos, COL.numero_solicitacao_coleta, S.Pontos_Coleta_idPontos_Coleta "&_
				"from Solicitacao_coleta COL "&_
				"inner join Solicitacao_coleta_has_Pontos_coleta S ON S.solicitacao_coleta_idsolicitacao_coleta = COL.idsolicitacao_coleta "&_
				"inner join Pontos_coleta PC ON PC.idpontos_coleta = S.pontos_coleta_idpontos_coleta "&_
				"WHERE COL.idSolicitacao_coleta = " & lIdSolicitacaoColeta					 
			   
		'response.write ssql
		Call search(sSql, arrCarta, intCarta)

		If intCarta > -1 Then
			For i=0 To intCarta
				lNrDoc = arrCarta(0,i)
				sRemetente = arrCarta(1,i)
				sCNPJ = arrCarta(2,i)
				sEndereco = arrCarta(3,i) & ", " & arrCarta(5,i) & " " & arrCarta(4,i)
				sCidade = arrCarta(6,i)
				sCEP = arrCarta(7,i)
				sUF = arrCarta(8,i)
				lQtd = arrCarta(9,i)			
				sNrSolCol = arrCarta(10,i)
				'lTipoColeta = arrCarta(11,i)
				lIdPontoColeta = arrCarta(11,i)
			Next
		Else
			sTxt = "Nenhum registro cadastrado."
		End If		
	End Sub
	
if sAcao = "1" then 'imprime a carta
%>
<html>
<body class="textoHome">
<div align="right">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="right"><a href="JavaScript:window.print()"><img src="../img/botao_print.gif" width="165" height="24" border="0"></a></div></td>
    </tr>
  </table>
  <br>
</div>
<table width=90% border="0" align="center" cellpadding="0" cellspacing="0" class="imgexpandeinfo">
		<tr>
			
    <td width="69%" height="25" id="explaintitle"><div align="left">&nbsp;<font size="2">Controle 
        para Remessa de Cartuchos Usados</font></div></td>
			
    <td width="31%" id="explaintitle"><div align="right">Documento N&ordm;&nbsp;&nbsp;<%=sNrSolCol%>&nbsp;&nbsp;</div></td>
		</tr>
		<tr>
			<td colspan=2>&nbsp;</td>
		</tr>
	</table>

	
<table width=90% border=0 align=center cellpadding="3" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td width="18%"><strong>Remetente:</strong></td>
    <td width="82%"><%=sRemetente%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td><strong>CNPJ/CPF:</strong></td>
    <td><%=sCNPJ%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td><strong>Endere�o:</strong></td>
    <td><%=sEndereco%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td><strong>Cidade:</strong></td>
    <td><%=sCidade%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td><strong>CEP:</strong></td>
    <td><%=sCEP%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textoHome"> 
    <td><strong>UF:</strong></td>
    <td><%=sUF%></td>
  </tr>
</table>
<%if lTipoColeta = 1 then%>
<p class="textoHome">Estamos remetendo � Okidata do Brasil Inform�tica Ltda o 
  material inserv�vel abaixo descriminado, <br />
  para fins de reciclagem e destina��o ambientalmente correta: <br />
<p>
<%else%>
<p class="textoHome">Material inserv�vel abaixo descriminado, <br />
  para fins de reciclagem e destina��o ambientalmente correta: <br />
<p>
<%end if%>

	
<TABLE width=90% border=0 align=center cellpadding="3" cellspacing="1" bgcolor="#999999">
  <TR bgcolor="#CCCCCC" class="textoHome"> 
    <TD width="23%"><font color="#000000"><strong>Descri��o</strong></font></TD>
    <TD width="33%"><font color="#000000"><strong>Qtd</strong></font></TD>
    <TD width="23%"><font color="#000000"><strong>Valor</strong></font></TD>
    <TD width="21%"><font color="#000000"><strong>Peso</strong></font></TD>
  </TR>
  <TR bgcolor="#FFFFFF" class="textoHome"> 
    <TD><strong>Cartuchos inserv�veis</strong></TD>
    <TD><%=lQtd%></TD>
    <TD>R$ <%=replace(formatnumber(lQtd, 2), ".", ",")%></TD>
    <TD><%=lQtd * 0.3%>Kg</TD>
  </TR>
</TABLE>
<p> <br />
  *Valor simb�lico. Remessa somente para tr�nsito, sem valor comercial.<br>
  **Peso aproximado total, considerando 0,3Kg por cartucho inserv�vel.<br />
<p>&nbsp;
<table width="90%" border="0" cellpadding="0" cellspacing="0" class="textoHome">
  <tr> 
	<%if 1 = 1 then%>
		  <td width="300" height="25" class="fonteMenu">Destinat�rio: </td>
		  <td width="300" class="fonteMenu">Local de entrega do material:</td>
		</tr>
		<tr> 
		  <td width="300" valign="top">Okidata do Brasil Informatica Ltda<br />
		    Endere�o: Avenida Alfredo Eg�dio de Souza Aranha,<br />
		    N&ordm; 100 / 4� e 5� andar � bloco C<br />
		    Cidade: S�o Paulo<br />
		    Estado: SP<br />
		    CNPJ: 01.619.318/0001-18<br />
		    I.E: 114.977.252.116</td>
		  <td width="300" valign="top">Atlas Log�stica Ltda.<br />
		    Endere�o: Avenida Aruan� 884<br />
		    Cidade: Barueri<br />
		    Estado: SP<br />
		    CNPJ:00.493.606/0001-06<br />
		    I.E: 206.076.757.110</td>
		</tr>
  <%
  else
		Response.Write "<td width=300 height=25 class=fonteMenu>Endere�o de entrega: </td>"
		Response.Write "</tr>"

		sSql = "select razao_social, logradouro, numero_endereco, complemento_endereco, municipio, estado, cnpj from Pontos_Coleta where IdPontos_Coleta = " & lIdPontoColeta
		
		Call search(sSql, arrCarta, intCarta)
		
		Response.Write "<td width=300 valign=top>"
		If intCarta > -1 Then
			For i=0 To intCarta
				Response.Write arrCarta(0,i) & "<br>"
				Response.Write "Endere�o: " & arrCarta(1,i) & ", N&ordm; " & arrCarta(2,i) & " " & arrCarta(3,i) & "<br>"
				Response.Write "Cidade: " & arrCarta(4,i) & "<br>"
				Response.Write "Estado: " & arrCarta(5,i) & "<br>"
				Response.Write "I.E: " & arrCarta(6,i)
			Next
		End If
		Response.Write "</td>"
  end if
  %>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p><table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="14%" class="textoHome">Local e Data:</td>
    <td width="41%">___________________________</td>
    <td width="9%" class="textoHome">Assinatura:</td>
    <td width="36%">_________________________</td>
  </tr>
</table>
<p>&nbsp;</p>
<p><br>
</body>
</html>
<%	
else	
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
			<title>
				<%=TITLE%>
			</title>
<link href="css/geral.css" rel="stylesheet" type="text/css"/>
<SCRIPT LANGUAGE="JavaScript">
function VoltaHome()
{
	<%if lAdm = 1 then%>
		window.location.href='frmAddSolicitacao.asp'
	<%else%>
		alert("<%if lTipoColeta = 0 then Response.Write "Seu cadastro foi submetido a aprova��o, em breve entraremos em contato para autorizar a entrega do(s) cartucho(s) no ponto de coleta!" else Response.Write "Seu cadastro foi submetido a aprova��o, em breve entraremos em contato para providenciar a coleta!" end if%>")
		window.location.href='index.asp?area=home'
	<%end if%>
}
</SCRIPT>

</head>
	<body>
		<script language="javascript" type="text/javascript" src="js/frmCadCliente.js"></script>
<div id="container">
			<!--#include file="inc/i_header.asp" -->
			<div id="conteudo">
				<table cellspacing="0" cellpadding="0" width="775">
					<tr>
						<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
						<td id="conteudo">
							<table cellpadding="3" cellspacing="4" width="100%" id="tableCadClienteCartaNF" style="display:;">
							<tr>
								<td align="right">
								<a align="right" class="linkoperacional" style="cursor:pointer;" onClick="VoltaHome();">
								<%if lAdm = 1 then%>
									Finalizar Solicita��o
								<%else%>
									Finalizar cadastro
								<%end if%>
								</a>
							</td>
							</tr>
								<tr>
									<td colspan="3" id="explaintitle">Carta de Remessa <%if lTipoColeta = 0 then Response.Write "(Ponto de Coleta)" else Response.Write "(Coleta Domiciliar)" end if%> / Nota Fiscal</td>
								</tr>
								<tr>
									<td colspan="3">
										<%if lTipoPessoa = 0 then%>
											<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
											 <tr>
												<td><font color="#666666" size="2" face="Verdana, Arial, Helvetica, sans-serif">Prezado(a)s 
												  cliente(s); </font><font face="Verdana, Arial, Helvetica, sans-serif"><p><font color="#666666" size="2">
												  <%if lTipoColeta = 1 then%>
														Para que a OKI do Brasil possa efetuar a coleta do material
												  <%else%>
														Para efetuar a entrega do material no ponto de coleta
												  <%end if%>
													, ser&aacute; obrigat&oacute;rio a apresenta&ccedil;&atilde;o 
													da Carta de Remessa impressa, assinada pelo respons&aacute;vel desta opera&ccedil;&atilde;o. 
													</font></p>
												  <p><font color="#666666" size="2">Esta Carta se trata de um documento de controle interno 
													da OKI Printins Solutions, no intuito de identificar o material provindo 
													do determinado cliente. <br>
													Nesta opera&ccedil;&atilde;o todo o transporte desde a coleta e posterior 
													armazenamento, estar&aacute; sob total responsabilidade da OKI do Brasil.</font></p>
												  <p><font color="#666666" size="2">Grato pela compreens&atilde;o.</font></p>
												  <p><font color="#666666" size="2">OKI do Brasil</font></p>
												  </font></td>
											  </tr>
											</table>									
										<%else%>
											<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
											 <tr>
												<td><font color="#666666" size="2" face="Verdana, Arial, Helvetica, sans-serif">Prezado(a)s 
												  cliente(s); </font><font face="Verdana, Arial, Helvetica, sans-serif"><p><font color="#666666" size="2">Para que a OKI do Brasil possa efetuar a coleta 
													do material, ser&aacute; obrigat&oacute;rio a apresenta&ccedil;&atilde;o 
													da Carta de Remessa impressa, assinada pelo respons&aacute;vel desta opera&ccedil;&atilde;o 
													</font></p>
												  <p><font color="#666666" size="2">Esta carta se trata de um documento de controle interno 
													da OKI Printins Solutions, no intuito de identificar o material provindo 
													do determinado cliente. <br>
													Nesta opera&ccedil;&atilde;o todo o transporte desde a coleta e posterior 
													armazenamento, estar&aacute; sob total responsabilidade da OKI do Brasil.</font></p>
												  <p><font color="#666666" size="2">Caso seja necess&aacute;rio, o modelo de Nota Fiscal de 
													Remessa deste material estar&aacute; dispon&iacute;vel para consulta. 
													A Nota Fiscal pode acompanhar a Carta de Remessa, que ser&aacute; retirada 
													junto ao material dispon&iacute;vel para a coleta.<br>
													<br>
													Grato pela compreens&atilde;o.</font></p>
												  <p><font color="#666666" size="2">OKI do Brasil</font></p>
												  </font></td>
											  </tr>
											</table>
										<%end if%>
									</td>
								</tr>
								<tr>
									<td colspan="3" width="100%">
										<div>
											<a class="linkoperacional" href="frmCartaDoacaoNF.asp?IdSolicitacaoColeta=<%=lIdSolicitacaoColeta%>&Acao=1&TipoPessoa=<%=lTipoPessoa%>" target="_blank">&nbsp;&nbsp;&nbsp;Clique Aqui para Visualizar a Carta de Remessa</a>
											<p>
											<%if lTipoPessoa = 1 then%>
											<a class="linkoperacional" href="modelos/Modelo_NF_OKI_Remessa.pdf" target="_blank">&nbsp;&nbsp;&nbsp;Clique Aqui para Visualizar o Modelo da Nota Fiscal</a>
											<%end if%>
										</div>
									</td>
								</tr>
							</table>
							</td>
						<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
					</tr>
				</table>
			</div>
		<!--#include file="../inc/i_bottom.asp" -->
	</body>
</html>
<%end if%>
<%Call close()%>
<!--#include file="../_config/colectobject.asp" -->
