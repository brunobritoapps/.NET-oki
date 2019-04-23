<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
    response.Expires = -1

     %>
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
              <table width="98%" border="0" align="center" cellpadding="2" cellspacing="3">
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - Cliente</div></td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/categoria.gif" width="32" height="32" alt="Cadastro de Categorias [Inserir Novas Categorias]" onClick="window.location.href='frmCategoriasAdm.asp'" /><br /> 
                    <a href="frmCategoriasAdm.asp" class="linkOperacional">Categorias</a> 
                  </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/cadcliente.gif" alt="Cadastros [Administração dos Cadastros de Clientes]" onClick="window.location.href='frmCadastroClienteAdm.asp'" /><br /> 
                    <a href="frmCadastroClienteAdm.asp" class="linkOperacional">Cadastro 
                    de Clientes</a> </td>
                  <td align="center" width="33%"> <img src="img/pontocoleta.gif" alt="Agrupamento [Administrar Grupos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmPontoColetaAdm.asp'" /><br /> 
                    <a href="frmPontoColetaAdm.asp" class="linkOperacional">Pontos 
                    de Coleta</a> </td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/contato.png" alt="Contatos [Administrar Contatos no Cliente]" onClick="window.location.href='frmContatosAdm.asp'" /><br /> 
                      <a href="frmContatosAdm.asp" class="linkOperacional">Usuários</a> 
                  </td>
                  <td align="center" width="33%"> <img src="img/cep_enderecos.gif" alt="CEP´s [Administrar Endereços]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmCepEndAdm.asp'" /><br /> 
                    <a href="frmCepEndAdm.asp" class="linkOperacional">CEP´s / 
                    Endereços</a> </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - Produtos</div></td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img src="img/produtos.gif" alt="Produtos [Administrar Produtos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmCadProdutosAdm.asp'" /><br /> 
                    <a href="frmCadProdutosAdm.asp" class="linkOperacional">Produtos</a> 
                  </td>
                  <td align="center" width="33%"> <img src="img/grupo_produtos.gif" alt="Grupo Produtos [Administrar Grupo Produtos]" height="38" align="absmiddle" class="imgexpandeinfo" onClick="window.location.href='frmGrupoProdutosAdm.asp'" /><br /> 
                    <a href="frmGrupoProdutosAdm.asp" class="linkOperacional">Grupo 
                    Produto</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Textos do Site" onClick="window.location.href='frmtiporelatorio.asp'" /><br /> 
                    <a href="frmtiporelatorio.asp" class="linkOperacional">Relatórios</a> 
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - Solicita&ccedil;&otilde;es</div></td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicitações de Coleta [Administrar Solicitações de Coleta]" onClick="window.location.href='frmSolicitacoesAdmMaster.asp'" /><br /> 
                    <a href="frmSolicitacoesAdmMaster.asp" class="linkOperacional">Solicitações 
                    de Coleta Master</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicitações de Coleta [Administrar Solicitações de Coleta]" onClick="window.location.href='frmSolicitacoesAdmDom.asp'" /><br /> 
                    <a href="frmSolicitacoesAdmDom.asp" class="linkOperacional">Solicitações 
                    de Coleta Domiciliar</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicitações de Coleta [Administrar Solicitações de Coleta]" onClick="window.location.href='frmSolicitacoesAdmPonto.asp'" /><br />
                    <a href="frmSolicitacoesAdmPonto.asp" class="linkOperacional">Solicitações 
                    de Coleta Ponto de Coleta</a> </td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Textos do Site" onClick="window.location.href='frmsolicitacoesresgateclientes.asp'" /><br /> 
                    <a href="frmsolicitacoesresgateclientes.asp" class="linkOperacional">Solicitação 
                    Resgate Clientes</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Kardex de Coleta [Administrar Recebimento de Coleta]" onClick="window.location.href='frmKardex_teste.asp'" /><br /> 
                    <a href="frmKardex_teste.asp" class="linkOperacional">Importar 
                    Arquivo de Solicitação</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Admnistração do Kardex" onClick="window.location.href='frmadmkardex.asp'" /><br /> 
                    <a href="frmadmkardex.asp" class="linkOperacional">Administração 
                    do Kardex</a> </td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Textos do Site" onClick="window.location.href='frmsolicitacoesresgatepontocoleta.asp'" /><br /> 
                    <a href="frmsolicitacoesresgatepontocoleta.asp" class="linkOperacional">Solicitação 
                    Resgate Ponto Coleta</a> </td>
                  <td align="center" width="33%"></td>
                  <td align="center" width="33%"></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - Transportadoras</div></td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/transportadoras.gif" height="40" alt="Transportadoras [Administração de Transportadoras]" onClick="window.location.href='frmTransportadorasAdm.asp'" /><br /> 
                    <a href="frmTransportadorasAdm.asp" class="linkOperacional">Cadastro 
                    de Transportadoras</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Arquivo eletrônico [Transportadora]" onClick="window.location.href='frmEletronicFileTransp.asp'" /><br /> 
                    <a href="frmEletronicFileTransp.asp" class="linkOperacional">Transportadora 
                    / Arquivo eletrônico</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Gerar Arquivo eletrônico [Transportadora]" onClick="window.location.href='frmGeraFileTransp.asp'" /><br /> 
                    <a href="frmGeraFileTransp.asp" class="linkOperacional">Transportadora 
                    / Gerar Arquivo</a> </td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/pasta_transportadora.jpg" width="35" height="35" alt="Arquivos Exportados [Transportadora]" onClick="window.location.href='frmlistafiletranspexport.asp'" /><br /> 
                    <a href="frmlistafiletranspexport.asp" class="linkOperacional">Transportadora 
                    / Listar Arquivos</a> </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - B&ocirc;nus</div></td>
                </tr>
                <tr> 
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Cadastro de Bônus [Administrar Bônus]" onClick="window.location.href='frmcadbonus.asp'" /><br /> 
                    <a href="frmcadbonus.asp" class="linkOperacional">Cadastro 
                    de Bônus</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Cadastro de Bônus [Administrar Bônus]" onClick="window.location.href='frmbonusgeradoclientes.asp'" /><br /> 
                    <a href="frmbonusgeradoclientes.asp" class="linkOperacional">Bônus 
                    Gerado [Cliente]</a> </td>
                  <td align="center" width="33%"> <img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Cadastro de Bônus [Administrar Bônus]" onClick="window.location.href='frmbonusgeradopontocoleta.asp'" /><br /> 
                    <a href="frmbonusgeradopontocoleta.asp" class="linkOperacional">Bônus 
                    Gerado [Ponto Coleta]</a> </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" id="explaintitle"><div align="center">Painel 
                      de Controle da Administração - Cadastros</div></td>
                </tr>
                <tr> 
				<td align="center" width="33%">
					<img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Cadastro de Banner" onClick="window.location.href='frmcadbanner.asp'" /><br />
					<a href="frmcadbanner.asp" class="linkOperacional">Cadastro de Banner</a>
				</td>
				<td align="center" width="33%">
					<img align="absmiddle" class="imgexpandeinfo" src="img/pasta_transportadora.jpg" width="35" height="35" alt="Cadastro de Notícias" onClick="window.location.href='frmcadnoticias.asp'" /><br />
					<a href="frmcadnoticias.asp" class="linkOperacional">Cadastro de Notícias</a>
				</td>
                <td align="center" width="33%">
					<img align="absmiddle" class="imgexpandeinfo" src="img/kardex.jpg" alt="Textos do Site" onClick="window.location.href='frmcadtextos.asp'" /><br />
                    <a href="frmcadtextos.asp" class="linkOperacional">Cadastro 
                    de Textos</a> </td>
                </tr>
                <tr> 
				<td align="center" width="33%">
					<img align="absmiddle" class="imgexpandeinfo" src="img/bonus.gif" width="35" height="35" alt="Upload de Imagem" onClick="window.location.href='frmuploadimagem.asp'" /><br />
					<a href="frmuploadimagem.asp" class="linkOperacional">Upload de Imagem</a>
				</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <div align="center"></div>
              
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
