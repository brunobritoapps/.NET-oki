<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="frmOperacionalAdm.asp" name="frmOperacionalAdm" method="POST">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<div id="painelcontrole">
						<table cellspacing="3" cellpadding="2" width="100%" border=0>
							<tr>
								<td colspan="3" id="explaintitle" align="center">Painel de Controle do Ponto de Coleta</td>
							</tr>
							<tr>
								<td id="explaintitle" align="right" colspan="3" style="padding:4px 4px 4px 4px;">
									<a href="?logoff=true" style="color: #FFFFFF;">Logoff do Sistema</a>
								</td>
							</tr>
							<!--tr>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/categoria.gif" width="32" height="32" alt="Cadastro de Categorias [Inserir Novas Categorias]" onClick="window.location.href='frmCategoriasAdm.asp'" /><br />
									<a href="frmCategoriasAdm.asp" class="linkOperacional">Categorias</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/contato.png" alt="Contatos [Administrar Contatos no Cliente]" onClick="window.location.href='frmContatosAdm.asp'" /><br />
									<a href="frmContatosAdm.asp" class="linkOperacional">Contatos</a>
								</td>
								<td align="center" width="33%">
									<img src="img/cep_enderecos.gif" alt="CEP�s [Administrar Endere�os]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmCepEndAdm.asp'" /><br />
									<a href="frmCepEndAdm.asp" class="linkOperacional">CEP�s / Endere�os</a>
								</td>
							</tr>
							<tr>
								<td align="center" width="33%">
									<img src="img/groupclient.png" alt="Agrupamento [Administrar Grupos]" height="39" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmAgrupamentosAdm.asp'" /><br />
									<a href="frmAgrupamentosAdm.asp" class="linkOperacional">Grupos</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/cadcliente.gif" alt="Cadastros [Administra��o dos Cadastros de Clientes]" onClick="window.location.href='frmCadastroClienteAdm.asp'" /><br />
									<a href="frmCadastroClienteAdm.asp" class="linkOperacional">Cadastro de Clientes</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/transportadoras.gif" height="40" alt="Transportadoras [Administra��o de Transportadoras]" onClick="window.location.href='frmTransportadorasAdm.asp'" /><br />
									<a href="frmTransportadorasAdm.asp" class="linkOperacional">Cadastro de Transportadoras</a>
								</td>
							</tr>
							<tr>
								<td align="center" width="33%">
									<img src="img/pontocoleta.gif" alt="Agrupamento [Administrar Grupos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmPontoColetaAdm.asp'" /><br />
									<a href="frmPontoColetaAdm.asp" class="linkOperacional">Pontos de Coleta</a>
								</td>
								<td align="center" width="33%">
									<img src="img/produtos.gif" alt="Produtos [Administrar Produtos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmCadProdutosAdm.asp'" /><br />
									<a href="frmCadProdutosAdm.asp" class="linkOperacional">Produtos</a>
								</td>
								<td align="center" width="33%">
									<img src="img/grupo_produtos.gif" alt="Grupo Produtos [Administrar Grupo Produtos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmGrupoProdutosAdm.asp'" /><br />
									<a href="frmGrupoProdutosAdm.asp" class="linkOperacional">Grupo Produto</a>
								</td>
							</tr-->
							<tr>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicita��es de Coleta [Solicita��es de Coleta Master]" onClick="window.location.href='frmacompanhasolicitacaomasterponto.asp'" /><br />
									<a href="frmacompanhasolicitacaomasterponto.asp" class="linkOperacional">Solicita��es de Coleta Master</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicita��es de Coleta [Baixa de Solicita��o de Coleta]" onClick="window.location.href='frmsolicitacoesentrega.asp'" /><br />
									<a href="frmsolicitacoesentrega.asp" class="linkOperacional">Baixa de Solicita��o de Coleta</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicita��es de Coleta [Acompanhamento de Solicita��o de Coleta]" onClick="window.location.href='frmacompanhasolicitacaoponto.asp'" /><br />
									<a href="frmacompanhasolicitacaoponto.asp" class="linkOperacional">Acompanhamento de Solicita��o de Coleta</a>
								</td>
							</tr>
							<tr>
							<% If not len(trim(session("CodBonus"))) = 0 Then %>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" width="35" height="35" src="img/bonus.gif" alt="B�nus Gerado" onClick="javascript:window.open('frmbonusgeradopontoadm.asp','','width=750,height=650,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /><br />
									<a href="frmbonusgeradopontoadm.asp" class="linkOperacional">B�nus Gerado</a>
								</td>
							<%End if%>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" width="35" height="35" src="img/kardex.jpg" alt="Relat�rios" onClick="window.location.href='frmtiporelatorioponto.asp';" /><br />
									<a href="frmtiporelatorioponto.asp" class="linkOperacional">Relat�rios</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicita��es de Coleta [Acompanhamento de Solicita��o de Coleta]" onClick="window.location.href='frmacompanhasolicitacaoresgateponto.asp'" /><br />
									<a href="frmacompanhasolicitacaoresgateponto.asp" class="linkOperacional">Solicita��es de Resgate</a>
								</td>
							</tr>
							<!--tr>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Kardex de Coleta [Administrar Recebimento de Coleta]" onClick="window.location.href='frmKardex.asp'" /><br />
									<a href="frmKardex.asp" class="linkOperacional">Importar Arquivo de Solicita��o</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Arquivo eletr�nico [Transportadora]" onClick="window.location.href='frmEletronicFileTransp.asp'" /><br />
									<a href="frmEletronicFileTransp.asp" class="linkOperacional">Transportadora / Arquivo eletr�nico</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Admnistra��o do Kardex" onClick="window.location.href='frmadmkardex.asp'" /><br />
									<a href="frmadmkardex.asp" class="linkOperacional">Admnistra��o do Kardex</a>
								</td>
							</tr>
							<tr>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Cadastro de B�nus [Administrar B�nus]" onClick="window.location.href='frmcadbonus.asp'" /><br />
									<a href="frmcadbonus.asp" class="linkOperacional">Cadastro de B�nus</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Gerar Arquivo eletr�nico [Transportadora]" onClick="window.location.href='frmGeraFileTransp.asp'" /><br />
									<a href="frmGeraFileTransp.asp" class="linkOperacional">Transportadora / Gerar Arquivo</a>
								</td>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/pasta_transportadora.jpg" width="35" height="35" alt="Arquivos Exportados [Transportadora]" onClick="window.location.href='frmlistafiletranspexport.asp'" /><br />
									<a href="frmlistafiletranspexport.asp" class="linkOperacional">Transportadora / Listar Arquivos</a>
								</td>
							</tr-->
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
