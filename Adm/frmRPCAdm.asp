<!--#include file="../_config/_config.asp" -->
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
								<td colspan="3" id="explaintitle" align="center">Painel de Controle da Administração</td>
							</tr>
							<tr>
							<tr>
								<td align="center" width="33%">
									<img align="absmiddle" class="imgexpandeinfo" src="img/solicitacao_new.png" alt="Solicitações de Coleta [Administrar Solicitações de Coleta]" onClick="window.location.href='frmSolicitacoesAdmPonto.asp'" /><br />
									<a href="frmSolicitacoesRPC.asp" class="linkOperacional">Solicitações de Resgate do Ponto de Coleta</a>

								</td>
								<td align="center" width="33%">
								</td>
								<td align="center" width="33%">
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
