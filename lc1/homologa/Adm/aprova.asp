
<% Call open() %>
<% Response.Charset="ISO-8859-1" %>

<%
    Dim sid

    sid = Request.QueryString("id")

        Dim objCDOSYSMail
        Dim objCDOSYSCon

        'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
        Set objCDOSYSMail = Server.CreateObject("CDO.Message")

        'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

        'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Oki!321!" '"!nfe321!"        'senha
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
        'objCDOSYSMail.To = "peterson.aquino@hotmail.com"
        'objCDOSYSMail.CC = "peterson.aquino@hotmail.com"

        'ASSUNTO DA MENSAGEM
        objCDOSYSMail.Subject = "Okidata - Sistema de Gerenciamento de Recolhimento de Suprimentos"

        'CONTEÚDO DA MENSAGEM
        objCDOSYSMail.HtmlBody = "Ola"

        'ENVIA A MENSAGEM
        objCDOSYSMail.Send

        'DESTRÓI OS OBJETOS
        Set objCDOSYSMail = Nothing
        Set objCDOSYSCon = Nothing

        return 'Ola'

     %>