<!--#include file="_config/_config.asp" -->
<!--#include file="inc/i_banner.asp" -->
<%Call open()%>
<% 
    Sub SubmitForm()
        if request.QueryString("cadastro")="ok" Then
            'response.Write "<script>alert('Cadastro efetuado com Sucesso! Por gentileza aguardar a aprovação do mesmo. Você receberá um e-mail de confirmação do cadastro e logo entraremos em contato!')</script>"

        End if    
        
    End SUb
    Call SubmitForm()

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<div id="container">
  <!--#include file="inc/i_header.asp" -->
  <div id="conteudo">
    <table cellspacing="0" cellpadding="0" width="775">
      <tr>
        <td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
        <td align="center" id="conteudo"><table width="750" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="510" valign="top"><br>
                <table width="480" border="0" align="center" cellpadding="3" cellspacing="3" class="textoHome">
                  <tr>
                    <td><div align="justify">
                        <%if len(trim(request.querystring("area"))) > 0 then%>
                        <%=getTextoByArea(request.querystring("area"))%>
                        <%elseif len(trim(request.querystring("idnoticia"))) > 0 then%>
                        <%=getNoticiasById(request.querystring("idnoticia"))%>
                        <%end if%>
                        <p>&nbsp;</p>
                      </div></td>
                  </tr>
				  <%
					if request.querystring("area") = "home" then
					%>
                  <tr>
                    <td><div align="center">
                        <table width="495" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><img src="img/img_topo_noticias.gif" width="495" height="27"></td>
                          </tr>
                          <tr>
                            <td valign="top" background="img/img_meio_noticias.gif"><table width="480" height="130" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td width="230" valign="top"><table width="230" border="0" cellspacing="0" cellpadding="3">
                                      <tr>
                                        <td class="fonteMenu">Noticias em Destaque:</td>
                                      </tr>
                                      <tr>
                                        <td><div class="textoHome">
                                            <marquee direction="up" scrolldelay="150" scrollamount="2" onMouseOver="" height="140px">
                                            <%=getNoticias()%>
                                            </marquee>
                                          </div></td>
                                      </tr>
                                    </table></td>
                                  <td><table width="230" border="0" align="right" cellpadding="3" cellspacing="0">
                                      <tr>
                                        <td><img src="img/tit_PreservaOki.gif" width="220" height="20"></td>
                                      </tr>
                                      <tr>
                                        <td valign="top"><br>
                                          <table width="230" border="0" align="center" cellpadding="0" cellspacing="0">
                                            <tr>
                                              <td width="160" class="fonteMenu">Obtenha 
                                                mais informa&ccedil;&otilde;es 
                                                sobre o Programa de Sustentabilidade 
                                                da Oki,<br>
                                                e sua rela&ccedil;&atilde;o com 
                                                as a&ccedil;&otilde;es <br>
                                                que visam a preserva&ccedil;&atilde;o 
                                                ambiental. <a class="fonteVermelho" href="index.asp?area=preserva"><br>
                                                saiba mais</a></td>
                                              <td width="70" align="center"><img src="img/logo_preserva.gif" width="68" height="91"></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table></td>
                          </tr>
                          <tr>
                            <td><img src="img/img_Base_noticias.gif" width="495" height="27"></td>
                          </tr>
                        </table>
                        <br>
                        <br>
                      </div></td>
                  </tr>
				  <%
					end if
					%>
                  <tr>
                    <td></td>
                  </tr>
                  <tr>
                    <td><div align="center">
                        <table width="200" border="0" align="center" cellpadding="1" cellspacing="1">
                          <%=getBannersRodape()%>
                        </table>
                      </div></td>
                  </tr>
                </table>
                <p align="center">&nbsp;</p></td>
              <td width="240" valign="top"><div align="center"><br>
                  <%=getLogin()%> 
                 </div>
                <table width="200" border="0" align="center" cellpadding="1" cellspacing="1">
                  <%=getBannersLateral()%>
                </table>
                <p align="center"><br>
                </p></td>
            </tr>
          </table></td>
        <td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
      </tr>
    </table>
  </div>
  <!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>