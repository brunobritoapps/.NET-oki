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
								<td colspan="3" id="explaintitle" align="center">Relatórios</td>
							</tr>
							<tr>
								<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalPonto.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td colspan="3">
									<ul>
										<li><a href="frmrelatoriosolicitacaocoletaponto.asp" class="linkOperacional">Relatório de Solicitação de Coleta</a></li>
										<li><a href="frmrelatoriosolicitacaocoletapontomaster.asp" class="linkOperacional">Relatório de Solicitação de Coleta Master</a></li>
										<!--<li><a href="frmrelatoriobonuspontocoletapontocoleta.asp" class="linkOperacional">Relatório de Bonus de Coleta Master</a></li>-->
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
