<!--#include file="../../_config/_config.asp" -->

<%Call open()%>
<% Response.Charset="ISO-8859-1" %>
<%

	If Request.QueryString("sub") = "getpontocoleta" Then
		Call GetPontoColeta()
	ElseIf Request.QueryString("sub") = "aprovarsolicitacao" Then
		Call AprovarSolColeta()
	ElseIf Request.QueryString("sub") = "atualizaridtranspsol" Then
		Call UpIDTranspDataProg()
	ElseIf Request.QueryString("sub") = "cnpjexists" Then
		Call CnpjExists()
	ElseIf Request.QueryString("sub") = "cnpjexists2" Then
		Call CnpjExists2()
	End If

	Sub GetPontoColeta()
		Dim sSql, arrPontoColeta, intPontoColeta, i
		Dim retorno
		retorno = ""
		sSql = "SELECT idPontos_coleta, nome_fantasia, cnpj FROM Pontos_coleta WHERE idPontos_coleta = " & Request.QueryString("value")
		Call search(sSql, arrPontoColeta, intPontoColeta)
		If intPontoColeta > -1 Then
			retorno = retorno & arrPontoColeta(0,0) & ";"
			retorno = retorno & arrPontoColeta(1,0) & ";"
			retorno = retorno & arrPontoColeta(2,0)
		End If
		Response.Write retorno
	End Sub

	Sub EnviarEmail(Destino)
		Dim MsgBody

		If bAprovado Then
			MsgBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
								"<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
								"<head> " & _
								"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
								"<title>Email Okidata</title> " & _
								"</head> " & _
								"<body> " & _
									"<div id=""container"" align=""center""> " & _
										"<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
											"<p>Prezado(a) cliente;<br /> " & _
											"Primeiramente agradecemos pela sua colaboração e ingresso ao OKI Eco Program, o Programa de Coleta e Destinação de Cartuchos OKI.<br /><br /><br /> " & _
											"A sua Solicitação de Coleta foi aprovada, acesse a sua interface pelo site (<a href=""http://200.225.91.166/sgrs/frmlogincliente.asp"">clique aqui</a>)<br /><br /> " & _
											"Caso tenha dúvidas, por favor fale conosco.<br />" & _
											"Grande São Paulo +55 (11) 3444-6747 <br />" & _
											"Demais localidades 0800-115577 <br />" & _
											"Horário de atendimento: Segunda a Sábado - das 8:00 às 20:00 <br />" & _
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
		End If

		'If Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "192" or Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "127" Then
		'	Dim EnviarMail

		'	Set EnviarMail = Server.CreateObject("CDONTS.NewMail")
		'	EnviarMail.From = "etanji@okidata.com.br"
		'	EnviarMail.Subject = "Sistema de Gerenciamento de retorno de suprimentos"
		'	EnviarMail.To = Destino
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
            objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.okidata.com.br"'"mail.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
		    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "nfe@okidata.com.br"'"sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
		    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!nfe321!"'"Oki7080! " '"!nfe321!"        'senha
            objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

			'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
			'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "174.37.245.58"

			'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
			'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
            objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

			'PORTA DO CDO
			'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

			'TEMPO DE TIMEOUT (EM SEGUNDOS)
			'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

			'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
			objCDOSYSCon.Fields.update
			Set objCDOSYSMail.Configuration = objCDOSYSCon

			'NOME DO REMETENTE, E-MAIL DO REMETENTE
			'objCDOSYSMail.From = "Etanji <etanji@okidata.com.br>"

			'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
    		objCDOSYSMail.From = "sustentabilidadeoki@sustentabilidadeoki.com.br"

		    ''''NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
		    objCDOSYSMail.To = Destino
            'objCDOSYSMail.CC = "peterson.aquino@hotmail.com"

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


    '
    'peterson aquino - 14-5-2014

	Sub EnviaEmailAtlas(Destino)

        Dim MsgBody
        Dim sql, arr, intarr, i
		Dim html, style

        html = ""
		sql = "select " & _
                "c.Clientes_idClientes, c.cep_coleta, c.logradouro_coleta, " & _
                "c.numero_endereco_coleta, c.comp_endereco_coleta, c.bairro_coleta, c.municipio_coleta, c.estado_coleta, " & _
                "c.contato_coleta, c.ddd_resp_coleta, c.telefone_resp_coleta, c.depto_resp_coleta, c.ramal_resp_coleta  " & _
                ",e.Transportadoras_idTransportadoras ,f.email ,e.razao_social, e.cnpj " & _
                ",a.numero_solicitacao_coleta " & _
                "from dbo.Solicitacao_coleta as a " & _
                "left outer join solicitacao_coleta_has_clientes as c on c.Solicitacao_coleta_idSolicitacao_coleta = a.idSolicitacao_coleta " & _
                "left outer join Solicitacao_coleta as d on a.idSolicitacao_coleta = d.idSolicitacao_coleta " & _
                "left outer join Clientes as e on e.idClientes = c.Clientes_idClientes " & _
                "left outer join Transportadoras as f on f.idTransportadoras = e.Transportadoras_idTransportadoras " & _
                "where a.idSolicitacao_coleta = " & Request.QueryString("id") & " "

        call search(sql, arr, intarr)

        if intarr > -1 then

            MsgBody = MsgBody & " "
            MsgBody = MsgBody & "<html xmlns=http://www.w3.org/1999/xhtml>  "
            MsgBody = MsgBody & "<head>  "
            MsgBody = MsgBody & "<meta http-equiv='Content-Type content=text/html;' charset='iso-8859-1' />  "
            MsgBody = MsgBody & "<title>Email Okidata</title>  "
            MsgBody = MsgBody & "</head>  "
            MsgBody = MsgBody & "<body>  "
            MsgBody = MsgBody & "<div id='container' align='center'>  "
            MsgBody = MsgBody & "<div id='conteudo' style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;' >  "
            MsgBody = MsgBody & "<p>Prezada Transportadora<br />  "
            MsgBody = MsgBody & "Informação Sobre Aprovação de Coleta.<br /><br /><br />  "
            MsgBody = MsgBody & "A Solicitação de coleta Id: <b>" & Request.QueryString("id") & " </b>Numero: <b>"  &  arr(17,0) & "</b> , foi aprovada, acesse a sua interface pelo site (<a href='http://www.sustentabilidadeoki.com.br/lc/homologa/adm'>clique aqui</a>)<br /><br />  "
            MsgBody = MsgBody & "&nbsp;<br />  "
            MsgBody = MsgBody & "</div>"
            MsgBody = MsgBody & "<table align='left' width='100%'>"
            MsgBody = MsgBody & "<tr>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td style='background:#990000;' align='center' colspan='2'><font style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px; color:#ffffff'><b>Dados da Coleta</b></font></td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Solic.Coleta:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(17,0) & "&nbsp;</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Cliente:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(15,0) & " - CPF/CNPJ: " & arr(16,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>				"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>CEP:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(1,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>								"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Endereço:</td>"
            MsgBody = MsgBody & "<td width='55%'>"& arr(2,0) & "," & arr(3,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
	    	MsgBody = MsgBody & "</tr>								"

            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Complemento:</td>"
            MsgBody = MsgBody & "<td width='55%'>"& arr(4,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>	"

            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Bairro:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(5,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>								"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Municipio:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(6,0) & " Estado: " & arr(7,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>					       "
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Contato:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(8,0) & " Telefone: (" & arr(9,0) & ") " & arr(10,0) & " Ramal: " & arr(12,0) & " Departamento: " & arr(11,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>							"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>Email:</td>"
            MsgBody = MsgBody & "<td width='55%'>" & arr(14,0) & "</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>							"
            MsgBody = MsgBody & "<tr style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;'>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='07%' align='right'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='55%'>&nbsp;</td>"
            MsgBody = MsgBody & "<td width='10%'>&nbsp;</td>"
            MsgBody = MsgBody & "</tr>							"
            MsgBody = MsgBody & "</table>"

			MsgBody = MsgBody & "<br />" & _
                "<div id='container' align='center'>  " & _
                "<div id='conteudo' style='width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;' >  " & _
                "Caso tenha dúvidas, por favor fale conosco.<br /> " & _
                "Grande São Paulo +55 (11) 3444-6747 <br /> " & _
                "Demais localidades 0800-115577 <br /> " & _
                "Horário de atendimento: Segunda a Sábado - das 8:00 às 20:00 <br /> " & _
                "Atenciosamente;<br /> " & _
                "<b style=color:#990000>OKI Printing Solutions</b>  " & _
                "</div>  " & _
                "<div id='bottom' style='font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;width:748px;'>  " & _
                "<p><b>CONFIDENCIALIDADE DO CORREIO ELETRÔNICO</b>  " & _
                "Esta mensagem, incluindo seus anexos, pode conter informação confidencial  " & _
                "e/ou privilegiada. Caso você tenha recebido este e-mail por engano, não  " & _
                "utilize, copie ou divulgue as informações nele contidas. E, por favor, avise  " & _
                "imediatamente o remetente, respondendo ao e-mail, e em seguida apague-o.</p>  " & _
                "<p><b>DISCLAIMER</b>  " & _
                "This message, including its attachments, may contain confidential and/or  " & _
                "privileged information. If you received this email by mistake, do not use,  " & _
                "copy or disseminate any information here in contained. Please notify us  " & _
                "immediately by replying to the sender and then delete it.</p>  " & _
                "</div>  " & _
                "</div>  " & _
                "</body>  " & _
                "</html>	"
        Else
            MsgBody = " Arr(" & arr &  "); Erro query " & sql
        End if

        Dim objCDOSYSMail
        Dim objCDOSYSCon

        'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
        Set objCDOSYSMail = Server.CreateObject("CDO.Message")

        'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

        'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.okidata.com.br"'"mail.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "nfe@okidata.com.br"'"sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!nfe321!"'"Oki7080! " '"!nfe321!"        'senha
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

        'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

        'PORTA DO CDO
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

        'TEMPO DE TIMEOUT (EM SEGUNDOS)
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

        'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
        objCDOSYSCon.Fields.update

        Set objCDOSYSMail.Configuration = objCDOSYSCon

        'NOME DO REMETENTE, E-MAIL DO REMETENTE
        objCDOSYSMail.From = "sustentabilidadeoki@sustentabilidadeoki.com.br"

        'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
        objCDOSYSMail.To = Destino
        'objCDOSYSMail.CC = "peterson.aquino@hotmail.com"

        'ASSUNTO DA MENSAGEM
        objCDOSYSMail.Subject = "Okidata - Sistema de Gerenciamento de Recolhimento de Suprimentos"

        'PARA ENVIO DA MENSAGEM NO FORMATO HTML, ALTERE O TextBody PARA HtmlBody
        objCDOSYSMail.HtmlBody = MsgBody

        'ENVIA A MENSAGEM
        objCDOSYSMail.Send

        'DESTRÓI OS OBJETOS
        Set objCDOSYSMail = Nothing
        Set objCDOSYSCon = Nothing

	End Sub

    Sub AprovarSolColeta()
        'EnviaEmailAtlas("sac.divlog@atlastranslog.com.br")
        'EnviaEmailAtlas("peterson.qna@gmail.com")
        AprovarSolColetaOld()

    End Sub


	Sub AprovarSolColetaOld()

		Dim sSql
		Dim sql, arr, intarr, i
		Dim sqltransp
        Dim cTransp
        Dim Aarr
        Dim iIntArr
        Dim cEmailTransp

		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
			     "SET [Status_coleta_idStatus_coleta] = 2 " & _
				 ",[data_aprovacao] = GetDate() " & _
				 "WHERE [idSolicitacao_coleta] = " & Request.QueryString("id")

        Call exec(sSql)

		if request("Tipo") = "C" then
    		sql = "select " & _
	    			"c.iscoletaemail, " & _
				    "c.idtransportadoras " & _
				    "from solicitacao_coleta_has_clientes as a " & _
				    "left join clientes as b " & _
				    "on a.clientes_idclientes  = b.idclientes " & _
				    "left join transportadoras as c " & _
				    "on b.transportadoras_idtransportadoras = c.idtransportadoras " & _
				    "where a.typecolect = 1 and a.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("id")
		else
		    sql = "select c.iscoletaemail,  c.idtransportadoras " & _
					    "from solicitacao_coleta_has_pontos_coleta as a " & _
    					"left join pontos_coleta as b on a.pontos_coleta_idpontos_coleta  = b.idPontos_coleta" & _
    				    "left join transportadoras as c on b.idtransp = c.idtransportadoras " & _
					    "where a.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("id")
		end if

		call search(sql, arr, intarr)

        '
        'criado novo recurso de disparo de e-mails
        'peter 17-5-2014
        if intarr > -1 then
            cTransp = arr(1,0)
            sSql = "select top 1 email from Transportadoras where idtransportadoras = " & cTransp

            call search(sSql, Aarr, iIntArr)

            if iIntArr > -1 Then
                cEmailTransp = Aarr(0,0)
            Else
                'caso não tenha achado e-mail, salva um default
                'cEmailTransp = "sustentabilidadeoki@sustentabilidadeoki.com.br"
                'cEmailTransp = "peterson.aquino@hotmail.com"
            End if
        End if

		if intarr > -1 then
			for i=0 to intarr
				if arr(0,i) > 0 then
					sqltransp = "insert into solicitacao_coleta_has_transportadoras ( " & _
									"solicitacao_coleta_idsolicitacao_coleta, " & _
									"transportadoras_idtransportadoras, " & _
									"numero_reconhecimento_transportadora) " & _
									"values ( " & _
									""&Request.QueryString("id")&", " & _
									""&arr(1,i)&", " & _
									"'NULL' " & _
									")"
					call exec(sqltransp)
				end if
			next
		end if

        '
        'peterson alteração 14-5-2014 - envio de e-mail de aprovação para ATLAS:
        EnviaEmailAtlas(cEmailTransp)

    End Sub


	Sub UpIDTranspDataProg()
		Dim sSql
		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
				   "SET " & _
				   "[data_programada] = " & Request.QueryString("data") & " " & _
				   "WHERE [idSolicitacao_coleta] = " & Request.QueryString("id")
		Call exec(sSql)
		sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
					   "([Solicitacao_coleta_idSolicitacao_coleta] " & _
					   ",[Transportadoras_idTransportadoras] " & _
					   ",[numero_reconhecimento_transportadora]) " & _
				 "VALUES " & _
					   "(" &Request.QueryString("id")& " " & _
					   "," &Request.QueryString("idtransp")& " " & _
					   ",NULL)"
		Call exec(sSql)
		Response.Write "Solicitação atualizada com sucesso!"
	End Sub

	Sub CnpjExists()
		Dim sSql, arrCnpj, intCnpj
		sSql = "SELECT [idPontos_coleta] " & _
			   "FROM [marketingoki2].[dbo].[Pontos_coleta] WHERE [cnpj] = '" & Request.QueryString("id") & "'"
'		Response.Write sSql
		Call search(sSql, arrCnpj, intCnpj)
		If intCnpj > -1 Then
			Response.Write "true"
		Else
			Response.Write "false"
		End If
	End Sub

	Sub CnpjExists2()
		Dim sSql, arrCnpj, intCnpj
		sSql = "SELECT [idTransportadoras] " & _
			   "FROM [marketingoki2].[dbo].[Transportadoras] WHERE [cnpj] = '" & Request.QueryString("id") & "'"
		Call search(sSql, arrCnpj, intCnpj)
		If intCnpj > -1 Then
			Response.Write "true"
		Else
			Response.Write "false"
		End If
	End Sub
%>
<%Call close()%>
