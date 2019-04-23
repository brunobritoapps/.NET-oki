<%@ Page Language="C#" AutoEventWireup="true" CodeFile="cs/frmesqueci.aspx.cs" Inherits="_Default" %>
  
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title>Recuperar Senha</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link rel="stylesheet" type="text/css" href="css/normalize.css" />
		<link rel="stylesheet" type="text/css" href="css/demo.css" />
		<link rel="stylesheet" type="text/css" href="css/component.css" />
</head>


<body>
<div id="container">
<link href="../css/geral.css" rel="stylesheet" type="text/css"> 
<div id="header">
  <table cellspacing="0" cellpadding="0" width="775" border="0">
    <tr>
      <td width="20"><img src="img/corner_supEsq.gif" width="20" height="20"></td>
      <td background="img/Bg_topo.gif">&nbsp;</td>
      <td width="20"><img src="img/corner_supDir.gif" width="20" height="20"></td>
    </tr>
  </table>
  <table width="775" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
      <td><table width="754" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="204" rowspan="2"><img src="img/img_logo.gif" width="204" height="85"></td>
            <td width="550" height="33" class="fonteMenu"><div align="right" class="fonteMenu">
              <table width="370" height="17" border="0" align="right" cellpadding="0" cellspacing="0">
                <tr>
                  <td><div align="center"><a href="index.asp?area=home" class="fonteMenu">&nbsp; 
                      Home&nbsp;</a></div></td>
                  <td width="1" bgcolor="5D5D5D"><div align="center"><img src="../img/_spacer.gif" width="1" height="1"></div></td>
                  <td><div align="center"><a href="index.asp?area=regulamento" class="fonteMenu">&nbsp;Regulamento </a> </div></td>
                  <td width="1" bgcolor="5D5D5D"><div align="center"><img src="../img/_spacer.gif" width="1" height="1"></div></td>
                  <td><div align="center"><a href="faleConosco.asp" class="fonteMenu">&nbsp;Fale 
                      Conosco </a></div></td>
                  <td width="1" bgcolor="5D5D5D"><div align="center"><img src="../img/_spacer.gif" width="1" height="1"></div></td>
                </tr>
              </table>
              <img src="../img/_spacer.gif" width="1" height="33"></td>
          </tr>
          <tr>
			<form action="frmbusca.asp" name="frmBusca" method="post" >
            <td width="550" height="52" bgcolor="DE2418" class="fontepadraoBold"><div align="right">
                <table width="170" border="0" cellspacing="3" cellpadding="3">
                  <tr>
                    <td class="fontepadraoBold">Procurar:</td>
                    <td><input name="txtBusca" type="text" class="TextBox" size="12" style="width:100px"></td>
                    <td class="fontepadraoBold"><img src="img/icon_procura.gif" class="imgexpandeinfo" alt="Buscar" width="13" height="13" onclick="document.frmBusca.submit()" /></td>
                  </tr>
                </table>
                   </div></td>
          			</form></tr>
          <tr>
            <td height="52" colspan="2"><table width="754" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="img/img_mata1.gif" width="377" height="209"></td>
                  <td width="377" background="img/img_mata2.gif"><table width="210" border="0" align="right" cellpadding="0" cellspacing="3">
                      <tr>
                        <td width="25"><img src="img/icon_folha.gif" width="21" height="24"></td>
                        <td><div align="left"><a class="fontepadraoBold" href="index.asp?area=PoliticaAmb">Política 
                            Ambiental OKI</a></div></td>
                      </tr>
                      <tr>
                        <td width="25"><img src="img/icon_vermelho.gif" width="21" height="24"></td>
                        <td><div align="left"><a class="fontepadraoBold" href="index.asp?area=coleta">Programa 
                            de Coleta</a></div></td>
                      </tr>
                      <tr>
                        <td><img src="img/icon_recicle.gif" width="21" height="24"></td>
                        <td><a class="fontepadraoBold" href="index.asp?area=residuos">Destina&ccedil;&atilde;o de Res&iacute;duos</a></td>
                      </tr>
                      <tr>
                        <td width="25"><img src="img/icon_default.gif" width="21" height="24"></td>
                        <td><div align="left"><a href="http://www.okidata.com/port/html/nf/WhereToBuy.html" target="_blank" class="fontepadraoBold">Onde Comprar</a></div></td>
                      </tr>
                      <tr>
                        <td width="25"><img src="img/icon_default.gif" width="21" height="24"></td>
                        <td><div align="left"><a href="http://www.okidata.com/port/html/nf/Home.html" target="_blank" class="fontepadraoBold">OKI do Brasil</a></div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
    </tr>
  </table>
</div>

	
	<form id="form1" runat="server">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<tr>
			<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>

			<td id="conteudo">
				<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="100%" class="oki-h1-bold">
							&nbsp;<br>
							&nbsp;Recupera&ccedil;&atilde;o de Senha:<br>
							<tr><td class="oki-h2">
							&nbsp;Para recuperar a sua senha preencha o campo abaixo com seu e-mail.<br></td></tr>
						</td>
					</tr>
					<tr>
						<td width="100%">&nbsp;<br>
							<asp:TextBox ID="txtBemail" type="text" runat="server" class="login-esqueci" name="txtBemail" />
							<asp:Button ID="Button1" class="btnform" runat="server" Text="Enviar" OnClick="Button1_Click" />
							</td>
						</td>
					</tr>
					<tr>
						<td width="200">&nbsp;<br>
						</td>
					</tr>					
				</table>
			</td>
			<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
		</tr>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
	</form>
  </div>
</body>
</html>