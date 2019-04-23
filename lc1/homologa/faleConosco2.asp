<!--#include file="_config/_config.asp" -->
<!--#include file="inc/i_banner.asp" -->
<%Call open()%>
<%
	Sub EnviarEmailOki()
		Dim MsgBody

		MsgBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
							"<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
							"<head> " & _
							"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
							"<title>Email Okidata</title> " & _
							"</head> " & _
							"<body> " & _
								"<div id=""container"" align=""center""> " & _
									"<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
										"<div>Nome: "&request.Form("textfield3")&" </div><br /> " & _
										"<div>Email: "&request.Form("textfield32")&" </div><br />" & _
										"<div>Telefone: "&request.Form("textfield33")&" </div><br />" & _
										"<div>Assunto: "&request.Form("textfield34")&" </div><br />" & _
										"<div>Mensagem: "&request.Form("textarea")&" </div><br />" & _
									"</div> " & _
									"<div id=""bottom"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;width:748px;""> " & _
										"<p><b>CONFIDENCIALIDADE DO CORREIO ELETRÔNICO</b> " & _
										"Esta mensagem, incluindo seus anexos, pode conter informação confidencial " & _
										"e/ou privilegiada. Caso você tenha recebido este e-mail por engano, não " & _
										"utilize, copie ou divulgue as informações nele contidas. E, por favor, avise " & _
										"imediatamente o remetente, respondendo ao e-mail, e em seguida apague-o.</p> " & _
										"<p><b>DISCLAIMER</b> " & _
										"This message, including its attachments, may contain confidential and/or " & _
										"privileged information. If you received this email by mistake, do not use, " & _
										"copy or disseminate any information here in contained. Please notify us " & _
										"immediately by replying to the sender and then delete it.</p> " & _
									"</div> " & _
								"</div> " & _
							"</body> " & _
							"</html>"

		'If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "192" or Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
		'	Dim EnviarMail

		'	Set EnviarMail = Server.CreateObject("CDONTS.NewMail")
		'	EnviarMail.From = request.Form("textfield32")
		'	EnviarMail.Subject = "Sistema de Gerenciamento de retorno de suprimentos"
		'	EnviarMail.To = "sustentabilidadeoki@marketingoki.com.br"
		'	EnviarMail.Body = MsgBody
		'	EnviarMail.Importance = 1
		'	EnviarMail.BodyFormat = 0
		'	EnviarMail.MailFormat = 0
		'	EnviarMail.Send
		'	Set EnviarMail = Nothing
		'Else
			Dim objCDOSYSMail
			Dim objCDOSYSCon
			'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
			Set objCDOSYSMail = Server.CreateObject("CDO.Message")

			'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
			Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

			'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
                        'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.marketingoki.com.br"
			
                        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.marketingoki.com.br"

			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sustentabilidadeoki@marketingoki.com.br" 'Email
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Oki!321!"        'senha

			'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
                        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587

			'PORTA DO CDO
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1

			'TEMPO DE TIMEOUT (EM SEGUNDOS)
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

			'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
			objCDOSYSCon.Fields.update
			Set objCDOSYSMail.Configuration = objCDOSYSCon

			'NOME DO REMETENTE, E-MAIL DO REMETENTE
			objCDOSYSMail.From = request.Form("textfield32")

			'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
			objCDOSYSMail.To = "sustentabilidadeoki@marketingoki.com.br"

			'ASSUNTO DA MENSAGEM
			objCDOSYSMail.Subject = "Okidata - Sistema de Gerenciamento de Recolhimento de Suprimentos"

			'CONTEÚDO DA MENSAGEM
			'objCDOSYSMail.TextBody = "Teste do componente CDOSYS"
			'PARA ENVIO DA MENSAGEM NO FORMATO HTML, ALTERE O TextBody PARA HtmlBody
			objCDOSYSMail.HtmlBody = MsgBody

			'ENVIA A MENSAGEM
			objCDOSYSMail.Send

			'DESTRÓI OS OBJETOS
			Set objCDOSYSMail = Nothing
			Set objCDOSYSCon = Nothing
		'End If
	End Sub

	Sub EnviarEmail()
		Dim MsgBody

		MsgBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
							"<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
							"<head> " & _
							"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
							"<title>Email Okidata</title> " & _
							"</head> " & _
							"<body> " & _
								"<div id=""container"" align=""center""> " & _
									"<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
										"<p>Prezado(a)<br /> " & _
										"Agradecemos primeiramente pelo contato e a atenção prestada.<br /><br /><br /> " & _
										"Em breve responderemos ao seu questionamento.<br /><br /> " & _
										"Atenciosamente;<br />" & _
										"<b style=""color:#990000"">OKI Printing Solutions</b> " & _
									"</div> " & _
									"<div id=""bottom"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;width:748px;""> " & _
										"<p><b>CONFIDENCIALIDADE DO CORREIO ELETRÔNICO</b> " & _
										"Esta mensagem, incluindo seus anexos, pode conter informação confidencial " & _
										"e/ou privilegiada. Caso você tenha recebido este e-mail por engano, não " & _
										"utilize, copie ou divulgue as informações nele contidas. E, por favor, avise " & _
										"imediatamente o remetente, respondendo ao e-mail, e em seguida apague-o.</p> " & _
										"<p><b>DISCLAIMER</b> " & _
										"This message, including its attachments, may contain confidential and/or " & _
										"privileged information. If you received this email by mistake, do not use, " & _
										"copy or disseminate any information here in contained. Please notify us " & _
										"immediately by replying to the sender and then delete it.</p> " & _
									"</div> " & _
								"</div> " & _
							"</body> " & _
							"</html>"

		If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "192" or Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
			Dim EnviarMail

			Set EnviarMail = Server.CreateObject("CDONTS.NewMail")
			EnviarMail.From = "sustentabilidadeoki@okidata.com.br"
			EnviarMail.Subject = "Sistema de Gerenciamento de retorno de suprimentos"
			EnviarMail.To = request.Form("textfield32")
			EnviarMail.Body = MsgBody
			EnviarMail.Importance = 1
			EnviarMail.BodyFormat = 0
			EnviarMail.MailFormat = 0
			EnviarMail.Send
			Set EnviarMail = Nothing
		Else
			Dim objCDOSYSMail
			Dim objCDOSYSCon
			'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
			Set objCDOSYSMail = Server.CreateObject("CDO.Message")

			'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
			Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

			'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
			'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
			
                        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "pop.okidata.com.br"

			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "nfe@okidata.com.br" 'Email
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!nfe321!"        'senha

			'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
                        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

			'PORTA DO CDO
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1

			'TEMPO DE TIMEOUT (EM SEGUNDOS)
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

			'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
			objCDOSYSCon.Fields.update
			Set objCDOSYSMail.Configuration = objCDOSYSCon

			'NOME DO REMETENTE, E-MAIL DO REMETENTE
			objCDOSYSMail.From = "sustentabilidadeoki@okidata.com.br"

			'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
			objCDOSYSMail.To = request("textfield32")

			'ASSUNTO DA MENSAGEM
			objCDOSYSMail.Subject = "Okidata - Sistema de Gerenciamento de Recolhimento de Suprimentos"

			'CONTEÚDO DA MENSAGEM
			'objCDOSYSMail.TextBody = "Teste do componente CDOSYS"
			'PARA ENVIO DA MENSAGEM NO FORMATO HTML, ALTERE O TextBody PARA HtmlBody
			objCDOSYSMail.HtmlBody = MsgBody

			'ENVIA A MENSAGEM
			objCDOSYSMail.Send

			'DESTRÓI OS OBJETOS
			Set objCDOSYSMail = Nothing
			Set objCDOSYSCon = Nothing
		End If
	End Sub

	if request.ServerVariables("HTTP_METHOD") = "POST" then
		'response.write "textfield3"&request("textfield3")&"<br />"
		'response.write "textfield32"&request("textfield32")&"<br />"
		'response.write "textfield33"&request("textfield33")&"<br />"
		'response.write "textfield34"&request("textfield34")&"<br />"
		'response.write "textarea"&request("textarea")&"<br />"
		'response.end
		call EnviarEmailOki()
		call EnviarEmail()
	end if
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
		<form action="faleconosco.asp" name="faleconosco" method="POST">
        <td align="center" id="conteudo"> <table width="750" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="510" valign="top"><br>
                <table width="480" border="0" align="center" cellpadding="3" cellspacing="3" class="textoHome">
                  <tr>
                    <td valign="top"> <div align="justify">
                        <table width="480" border="0" cellpadding="2" cellspacing="2" class="fonteMenu">
                          <tr>
                            <td><img src="img/tit_FaleConosco.gif" width="128" height="16"></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td>Nome:</td>
                          </tr>
                          <tr>
                            <td><input name="textfield3" type="text" class="TextBox" /></td>
                          </tr>
                          <tr>
                            <td>E-mail:</td>
                          </tr>
                          <tr>
                            <td><input name="textfield32" type="text" class="TextBox" /></td>
                          </tr>
                          <tr>
                            <td>Telefone:</td>
                          </tr>
                          <tr>
                            <td><input name="textfield33" type="text" class="TextBox" /></td>
                          </tr>
                          <tr>
                            <td>Assunto:</td>
                          </tr>
                          <tr>
                            <td><input name="textfield34" type="text" class="TextBox" /></td>
                          </tr>
                          <tr>
                            <td>Mensagem:</td>
                          </tr>
                          <tr>
                            <td><textarea name="textarea" rows="8" class="TextBox"></textarea></td>
                          </tr>
                          <tr>
                            <td><div align="center"><img src="img/botao_enviar_box.gif" width="47" height="19" onclick="document.faleconosco.submit()"><!--input type="submit" name="btnenviar" value="Enviar" /--></div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="50"><div align="center">(11) 3444-6747
                                - Grande S&atilde;o Paulo<br>
                                0800115577 - demais localidades<br>
                                Hor&aacute;rio de atendimento: Segunda a S&aacute;bado
                                - das 8:00 &agrave;s 20:00.</div></td>
                          </tr>
                        </table>
						</form>
                        <p>&nbsp;</p>
                      </div></td>
                  </tr>
                </table>
                <p align="center">&nbsp;</p></td>
              <td width="240" valign="top"><br>
				<%=getLogin()%>
                <p>&nbsp;</p>
                <table width="200" border="0" align="center" cellpadding="1" cellspacing="1">
                  <tr>
                    <th scope="col">&nbsp;</th>
                  </tr>
                  <tr>
                    <th scope="col">&nbsp;</th>
                  </tr>
                  <tr>
                    <th scope="col">&nbsp;</th>
                  </tr>
                </table>                <p align="center"><br>
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
