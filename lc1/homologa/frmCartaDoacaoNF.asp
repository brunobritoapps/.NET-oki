<!--#include file="_config/_config.asp" -->
<link href="css/geral.css" rel="stylesheet" type="text/css"/>
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
	Dim sTransId, sTransCnpj, sTransEnd, sTransCid, sTransUF, sTransIE, sTRNome, sTransComp
	Dim sContato, sTelefone, sRamal, sDepto
	Dim lIdPontoColeta
	Dim lAdm
	
	lTipoColeta = 1
	sAcao = request("Acao")
	lIdSolicitacaoColeta = request("IdSolicitacaoColeta")
	lTipoPessoa = request("TipoPessoa")
	lAdm = request("Adm")
	
	'============================================================================================

	if sAcao = "1" then
		call geraCartaDoacao()
	else
		'lTipoColeta	= request("TipoColeta")
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
					 
		'sSql = "SELECT S.Solicitacao_coleta_idSolicitacao_Coleta, C.Razao_Social, C.CNPJ, "&_
		'			"CEP.Logradouro, C.Compl_Endereco_coleta, C.Numero_Endereco_coleta, "&_
		'			 "CEP.Municipio, CEP.cep, CEP.Estado, COL.qtd_cartuchos, COL.numero_solicitacao_coleta, C.typeColect, S.Pontos_Coleta_idPontos_Coleta from clientes C "&_
		'			 "INNER JOIN Solicitacao_coleta_has_Clientes S ON C.IdClientes = S.Clientes_IdClientes "&_
		'			 "INNER JOIN Solicitacao_coleta COL ON S.Solicitacao_coleta_idSolicitacao_coleta = COL.idSolicitacao_coleta "&_
		'			 "INNER JOIN cep_consulta_has_Clientes CEP ON C.IdClientes = CEP.Clientes_IdClientes "&_
		'			 "WHERE COL.idSolicitacao_coleta = " & lIdSolicitacaoColeta					 
		
	If Len(Trim(lIdSolicitacaoColeta)) = 12 Then
		sSql = "SELECT SC.Solicitacao_coleta_idSolicitacao_Coleta, C.Razao_Social, C.CNPJ,"&_
			"SC.logradouro_coleta, SC.comp_endereco_coleta, SC.numero_endereco_coleta,"&_
			"SC.municipio_coleta, SC.cep_coleta, SC.estado_coleta,"&_
			"COL.qtd_cartuchos,COL.numero_solicitacao_coleta, C.typeColect, SC.Pontos_Coleta_idPontos_Coleta,"&_
			"SC.contato_coleta, SC.ddd_resp_coleta, SC.telefone_resp_coleta, SC.ramal_resp_coleta, SC.depto_resp_coleta, "&_			
			"ISNULL(TRANSC.cnpj,'N') AS CGC, TRANSC.razao_social NOME, "&_ 
            "ENDE.logradouro+', '+CAST(TRANSC.numero_endereco as varchar(10)) ENDERECO, ENDE.municipio MUNICIPIO, "&_ 
            "ENDE.estado UF, TRANSC.inscricao_estadual INSCM, TRANSC.compl_endereco COMP "&_
			"from clientes C "&_
			" INNER JOIN Solicitacao_coleta_has_Clientes SC ON C.IdClientes = SC.Clientes_IdClientes "&_
			" INNER JOIN Solicitacao_coleta COL ON SC.Solicitacao_coleta_idSolicitacao_coleta = COL.idSolicitacao_coleta "&_
			" LEFT JOIN Transportadoras TRANS ON TRANS.idTransportadoras = C.Transportadoras_idTransportadoras "&_
            " LEFT JOIN Clientes TRANSC ON TRANS.cnpj = TRANSC.cnpj "&_
            " LEFT JOIN cep_consulta_has_Clientes ENDE ON ENDE.Clientes_idClientes = TRANSC.idClientes "&_
			" WHERE COL.numero_solicitacao_coleta = '" & lIdSolicitacaoColeta & "'"
	Else
		sSql = "SELECT SC.Solicitacao_coleta_idSolicitacao_Coleta, C.Razao_Social, C.CNPJ,"&_
			"SC.logradouro_coleta, SC.comp_endereco_coleta, SC.numero_endereco_coleta,"&_
			"SC.municipio_coleta, SC.cep_coleta, SC.estado_coleta,"&_
			"COL.qtd_cartuchos,COL.numero_solicitacao_coleta, C.typeColect, SC.Pontos_Coleta_idPontos_Coleta,"&_
			"SC.contato_coleta, SC.ddd_resp_coleta, SC.telefone_resp_coleta, SC.ramal_resp_coleta, SC.depto_resp_coleta, "&_			
			"ISNULL(TRANSC.cnpj,'N') AS CGC, TRANSC.razao_social NOME, "&_ 
            "ENDE.logradouro+', '+CAST(TRANSC.numero_endereco as varchar(10)) ENDERECO, ENDE.municipio MUNICIPIO, "&_ 
            "ENDE.estado UF, TRANSC.inscricao_estadual INSCM, TRANSC.compl_endereco COMP "&_
			"from clientes C "&_
			" INNER JOIN Solicitacao_coleta_has_Clientes SC ON C.IdClientes = SC.Clientes_IdClientes "&_
			" INNER JOIN Solicitacao_coleta COL ON SC.Solicitacao_coleta_idSolicitacao_coleta = COL.idSolicitacao_coleta "&_
			" LEFT JOIN Transportadoras TRANS ON TRANS.idTransportadoras = C.Transportadoras_idTransportadoras "&_
            " LEFT JOIN Clientes TRANSC ON TRANS.cnpj = TRANSC.cnpj "&_
            " LEFT JOIN cep_consulta_has_Clientes ENDE ON ENDE.Clientes_idClientes = TRANSC.idClientes "&_
			" WHERE SC.Solicitacao_coleta_idSolicitacao_Coleta = '" & lIdSolicitacaoColeta & "'"
	End If
            '" WHERE COL.numero_solicitacao_coleta = '" & lIdSolicitacaoColeta & "' "
		'response.write "<td>" & ssql
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
				lTipoColeta = arrCarta(11,i)
				lIdPontoColeta = arrCarta(12,i)
				
				sContato = arrCarta(13,i)
				sTelefone = "(" & arrCarta(14,i) & ") " & arrCarta(15,i)
				sRamal = arrCarta(16,i)
				sDepto = arrCarta(17,i)
				
				sTransCnpj  = arrCarta(18,i)
				sTRNome     = arrCarta(19,i)
				sTransEnd   = arrCarta(20,i)
				sTransCid   = arrCarta(21,i)
				sTransUF    = arrCarta(22,i)
				sTransIE    = arrCarta(23,i)
				sTransComp  = arrCarta(24,i)
			Next
		Else
			sTxt = "Nenhum registro cadastrado."
		End If		
	End Sub
	
if sAcao = "1" then 'imprime a carta
%>
<html>
<style type="text/css">
<!--
div.TipoPequeno {mso-style-name:"Tipo Pequeno";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:6.0pt;
	margin-left:0cm;
	line-height:110%;
	font-size:7.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	letter-spacing:.2pt;}
li.TipoPequeno {mso-style-name:"Tipo Pequeno";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:6.0pt;
	margin-left:0cm;
	line-height:110%;
	font-size:7.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	letter-spacing:.2pt;}
p.TipoPequeno {mso-style-name:"Tipo Pequeno";
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:6.0pt;
	margin-left:0cm;
	line-height:110%;
	font-size:7.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	letter-spacing:.2pt;}
div.MsoNormal {margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;}
div.Obrigado {mso-style-name:Obrigado;
	margin-top:3.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:-7.1pt;
	margin-bottom:.0001pt;
	text-align:justify;
	font-size:8.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	font-style:italic;}
li.MsoNormal {margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;}
li.Obrigado {mso-style-name:Obrigado;
	margin-top:3.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:-7.1pt;
	margin-bottom:.0001pt;
	text-align:justify;
	font-size:8.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	font-style:italic;}
p.MsoNormal {margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;}
p.Obrigado {mso-style-name:Obrigado;
	margin-top:3.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:-7.1pt;
	margin-bottom:.0001pt;
	text-align:justify;
	font-size:8.0pt;
	font-family:"Microsoft Sans Serif",sans-serif;
	color:gray;
	font-style:italic;}
-->
</style>
<body class="textocarta">
<div align="right">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="right"><a href="JavaScript:window.print()"><img src="img/botao_print.gif" width="165" height="24" border="0"></a></div></td>
    </tr>
  </table>
  <table width="130" border="0">
  <tr>
    <td width="657"><img src="http://ftpodb.okidata.com.br/okilogo/OKI.png" alt="Oki Data" width="115" height="29"></td>
    </tr>
</table>
  <h1 align=center style='text-align:center'><span style='font-size:16.0pt;
font-family:"Verdana",sans-serif'>PROGRAMA DE SUSTENTABILIDADE OKI</span></h1>
  <h1 align=center style='text-align:center'><span style='font-size:16.0pt;
font-family:"Verdana",sans-serif'>TERMO DECLARAT&Oacute;RIO DE DOA&Ccedil;&Atilde;O</span>  </h1>
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
  <tr bgcolor="#FFFFFF" class="textocarta"> 
    <td width="18%"><strong>Remetente:</strong></td>
    <td width="82%"><%=sRemetente%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textocarta"> 
    <td><strong>CNPJ/CPF:</strong></td>
    <td><%=sCNPJ%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textocarta"> 
    <td><strong>Endere�o:</strong></td>
    <td><%=sEndereco%></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="textocarta"> 
    <td><strong>Cidade:</strong></td>
    <td><%=sCidade%></td>
  </tr>
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>CEP:</strong></td>
		<td><%=sCEP%></td>
	</tr>
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>UF:</strong></td>
		<td><%=sUF%></td>
	</tr>
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>Respons�vel:</strong></td>
		<td><%=sContato%></td>
	</tr>  
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>Telefone:</strong></td>
		<td><%=sTelefone%>&nbsp;</td>
	</tr> 
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>Ramal:</strong></td>
		<td><%=sRamal%></td>
	</tr> 
	<tr bgcolor="#FFFFFF" class="textocarta"> 
		<td><strong>Departamento:</strong></td>
		<td><%=sDepto%></td>
	</tr>
</table>

<%if lTipoColeta = 1 then%>
<p class="textocarta">Estamos remetendo � Okidata do Brasil Inform�tica Ltda o 
  material inserv�vel abaixo descriminado, <br />
  para fins de reciclagem e destina��o ambientalmente correta: <br />
<p>
<%else%>
<p class="textocarta">Material inserv�vel abaixo descriminado, <br />
  para fins de reciclagem e destina��o ambientalmente correta: <br />
<p>
<%end if%>
<p><span style="font-size:7.5pt;font-family:&quot;Verdana&quot;,sans-serif">Em conson&acirc;ncia ao que
disp&otilde;e a Lei 12.305/2010, que instituiu a Pol&iacute;tica Nacional de Res&iacute;duos
S&oacute;lidos, o Decreto n&deg; 7.404/2010 e demais legisla&ccedil;&otilde;es ambientais aplic&aacute;veis, remetemos
&agrave; Oki Data do Brasil Inf. Ltda., a t&iacute;tulo de doa&ccedil;&atilde;o, o (s) produto(s) inserv&iacute;vel
(is) abaixo descriminado(s) para fins de destina&ccedil;&atilde;o final ambientalmente
adequada.&nbsp;</span>
<p>
<TABLE width=90% border=0 align=center cellpadding="3" cellspacing="1" bgcolor="#999999">
  <TR bgcolor="#CCCCCC" class="textocarta"> 
    <TD width="23%"><font color="#000000"><strong>Descri��o</strong></font></TD>
    <TD width="33%"><font color="#000000"><strong>Qtd</strong></font></TD>
    <TD width="23%"><font color="#000000"><strong>Valor</strong></font></TD>
    <TD width="21%"><font color="#000000"><strong>Peso</strong></font></TD>
  </TR>
  <TR bgcolor="#FFFFFF" class="textocarta"> 
    <TD><strong>Cartuchos inserv�veis</strong></TD>
    <TD><%=lQtd%></TD>
    <TD>R$ <%=replace(formatnumber(lQtd, 2), ".", ",")%></TD>
    <TD><%=lQtd * 0.3%>Kg</TD>
  </TR>
</TABLE>
<p>
  *Valor simb�lico. Remessa somente para tr�nsito, sem valor comercial.<br>
  **Peso aproximado total, considerando 0,3Kg por cartucho inserv�vel.
<p>

<table width="90%" border="0" cellpadding="0" cellspacing="0" class="textocarta">
  <tr> 
	<%if lTipoColeta = 1 then%>
		  <td width="300" height="25" class="fonteMenu">Destinat�rio: </td>
		  <td width="300" class="fonteMenu">Local de entrega do material:</td>
  </tr>
		<tr> 
		  <td width="300" valign="top">Okidata do Brasil Informatica Ltda<br />
		    Endere�o: Avenida Alfredo Eg�dio de Souza Aranha,<br />
		    N&ordm; 100 / 5� andar / bloco C<br />
		    Cidade: S�o Paulo<br />
		    Estado: SP<br />
		    CNPJ: 01.619.318/0001-18<br />
		    I.E: 114.977.252.116</td>
		  <td width="300" valign="top">
          <%If sTransCnpj = "N" then%>
		  <!--  Atlas Log�stica Ltda.<br />
		    Endere�o: Avenida Aruan� 884<br />
		    Comp.: <br />
		    Cidade: Barueri<br />
		    Estado: SP<br />
		    CNPJ: 00.493.606/0001-06<br />
		    I.E: 206.076.757.110</td>    -- ANDRE AGGIO 22/06/2015 ALTERADO PARA OPERA��O RESOLUTION -->
		    RESOLUTION ECOLOG�STICA REVERSA LTDA.<br />
		    Endere�o: Rua Marina Ciufuli Zanfelice, 280<br />
		    Comp.: <br />
		    Cidade: S�o Paulo<br />
		    Estado: SP<br />
		    CNPJ: 17.096.898/0001-46<br />
		    I.E: 140.563.674.118</td>
          <%Else%>
		    <%=sTRNome%><br />
		    Endere�o: <%=sTransEnd%><br />
		    Comp.: <%=sTransComp%><br />
		    Cidade: <%=sTransCid%><br />
		    Estado: <%=sTransUF%><br />
		    CNPJ: <%=sTransCnpj%><br />
		    I.E: <%=sTransIE%></td>
		  <%End If%>
		</tr>
  <%
  else
		Response.Write "<td width=300 height=25 class=fonteMenu>Endere�o de entrega: </td>"
		Response.Write "</tr>"

		'sSql = "select razao_social, logradouro, numero_endereco, complemento_endereco, municipio, estado, cnpj from Pontos_Coleta where IdPontos_Coleta = " & lIdPontoColeta
		
		Response.Write "<td>Acao: " & sAcao
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
    <td height="22">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="708" border="0">
  <tr>
    <td width="378"><span style="font-size:9.0pt;
line-height:110%;font-family:&quot;Verdana&quot;,sans-serif;color:windowtext">Envio do
  produto inserv&iacute;vel autorizado por:</span></td>
    <td width="320">________________________________________</td>
  </tr>
  <tr>
    <td><span style="font-size:9.0pt;line-height:110%;font-family:&quot;Verdana&quot;,sans-serif;
color:windowtext">Local e Data / Assinatura:</span></td>
    <td>________________________________________</td>
  </tr>
</table>
<p align="center"><span style="font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:black">Em caso
  de d&uacute;vidas, entre em contato no telefone (11) 3444-3500 / Ramal 3563.</span>
</p>
 <br>
<br>
<br>
<p class=TipoPequeno style='margin-left:7.1pt;text-align:justify;text-indent:
-7.1pt'><span style='font-size:9.0pt;line-height:110%;font-family:"Verdana",sans-serif;
color:windowtext'>Observa&ccedil;&atilde;o: Opera&ccedil;&atilde;o fora do campo de incid&ecirc;ncia do ICMS</span></p>
<p class=Obrigado style='margin-left:0cm'>As informa&ccedil;&otilde;es contidas neste termo
  relativas &agrave; descri&ccedil;&atilde;o dos produtos, volume, valor e peso s&atilde;o de
  responsabilidade exclusiva do doador, estando sujeitas a confer&ecirc;ncia pela OKI
  DATA do BRASIL INFORM&Aacute;TICA LTDA. ap&oacute;s o recebimento do (s) produto (s)
  inserv&iacute;vel(is). </p>
<p class=MsoNormal></p>
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
						<td width="11" background="img/Bg_LatEsq.gif">&nbsp;&nbsp;</td>
						<td id="conteudo">
							<table cellpadding="3" border="0" cellspacing="4" width="100%" id="tableCadClienteCartaNF" style="display:;">
							</tr>
								<tr>
									<td colspan="3" id="explaintitle" style="text-align:center;"><h3>IMPORTANTE ! Estamos quase finalizando !</h3>
                                        <h3>Sua solicita��o j� est� em an�lise e ap�s aprova��o entraremos em contato. Ainda falta a impress�o da Carta de Remessa / Nota Fiscal</h3></td>
								</tr>
								<tr>
									<td colspan="3">
										<%if lTipoPessoa = 0 then%>
											<table width="98%" align="center" border="0" cellspacing="0" cellpadding="0">
											 <tr>
												<td><font color="#666666" size="2" face="Verdana, Arial, Helvetica, sans-serif">Prezado(a)s 
												  cliente(s)<font face="Verdana, Arial, Helvetica, sans-serif"><p><font color="#666666" size="2">
												  <%if lTipoColeta = 1 then%>
														Para que a OKI do Brasil possa efetuar a coleta do material
												  <%else%>
														Para efetuar a entrega do material no ponto de coleta
												  <%end if%>
													, ser&aacute; obrigat&oacute;rio a apresenta&ccedil;&atilde;o 
													da Carta de Remessa impressa, assinada pelo respons&aacute;vel desta opera&ccedil;&atilde;o. 
													</font></p>
												  <p><font color="#666666" size="2" style="text-align: justify">Esta Carta se trata de um documento de controle interno 
													da OKI Printins Solutions, no intuito de identificar o material provindo 
													do determinado cliente. </font></p>
                                                    <p><font color="#666666" size="2" style="text-align: justify"> <br>
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
									<td colspan="1" width="50%" align="center">
										<div style="vertical-align: middle;">
											<a class="linkoperacional" href="frmCartaDoacaoNF.asp?IdSolicitacaoColeta=<%=lIdSolicitacaoColeta%>&Acao=1&TipoPessoa=<%=lTipoPessoa%>" target="_blank">Clique Aqui para Visualizar a Carta de Remessa</a>
										</div>
									</td>
									<td>&nbsp;</td>
									<td align="center">
										<a class="linkoperacional" style="cursor:pointer;" onClick="VoltaHome();">
                                        Finalizar
										</a>
									</td>	
								</tr>
							</table>
						</td>
						<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
					</tr>
				</table>
			</div>
		<!--#include file="inc/i_bottom.asp" -->
	</body>
</html>
<%end if%>
<%Call close()%>
<!--#include file="_config/colectobject.asp" -->
