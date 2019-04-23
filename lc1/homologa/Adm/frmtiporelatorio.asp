<!--#include file="../_config/_config.asp" -->
<%session ("sql")=""%>
<%Call open()%>
<%Call GetSessionAdm()%>
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
								<td colspan="3" id="explaintitle" align="center">Relatórios</td>
							</tr>
							<tr>
								<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td colspan="3">
									<ul>
										<li><a href="frmrelatoriocliente.asp" class="linkOperacional">Clientes</a></li>
										<li><a href="frmrelatoriosolicitacaocoletadomiciliar.asp" class="linkOperacional">Coletas</a></li>
										<li><a href="frmrelatoriobonuscliente.asp" class="linkOperacional">Bônus Cliente</a></li>
                                        <li><a href="frmrelatorioNovoResgateBonificacao.asp" class="linkOperacional">Resgate de Bonificações - novo</a></li>
									</ul>
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
