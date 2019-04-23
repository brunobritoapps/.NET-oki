<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<% Response.Charset="ISO-8859-1" %>
<%

	If Request.QueryString("sub") = "getpontocoleta" Then
		Call GetPontoColeta()
	ElseIf Request.QueryString("sub") = "aprovarsolicitacao" Then
		Call AprovarSolColeta()	
	ElseIf Request.QueryString("sub") = "aprovarsolicitacaocli" Then
		Call AprovarSolColetaCli()
	ElseIf Request.QueryString("sub") = "atualizaridtranspsol" Then
		Call UpIDTranspDataProg()
	ElseIf Request.QueryString("sub") = "cnpjexists" Then
		Call CnpjExists()
	ElseIf Request.QueryString("sub") = "cnpjexists2" Then
		Call CnpjExists2()		
	ElseIf Request.QueryString("sub") = "getprod" Then
		Call GetProds()
	End If

	Sub GetProds()
		Dim sSql, arrProds, intProds, i, n
		Dim retorno
		retorno = "<select name=Prods id=Prods multiple size=8 class=controls style='width: 330px; height: 210px; font-size: 13px;'>"
		
		if len(trim(request("IdGrupo"))) = 0 then
			sSql = "select * from dbo.Produtos where idoki LIKE '%" & request("SearchProd") & "%' OR descricao LIKE '%"& request("SearchProd") &"%'"
		elseif len(trim(request("IdGrupo"))) > 0 and len(trim(request("SearchProd"))) > 0 then
			sSql = "select * from dbo.Produtos where (idoki LIKE '%" & request("SearchProd") & "%' OR descricao LIKE '%"& request("SearchProd") &"%') AND grupo_produtos_idgrupo_produtos = " & request("IdGrupo")
		else
			sSql = "select * from dbo.Produtos where grupo_produtos_idgrupo_produtos = " & request("IdGrupo")
		end if
		
		Call search(sSql, arrProds, intProds)
		If intProds > -1 Then
			for n=0 to intProds
				retorno = retorno & "<option value="""& trim(arrProds(0,n)) &""">"& arrProds(2,n) &"</option>" & vbcrlf
			next
			retorno = retorno & "</select>"
		else
			retorno = retorno & "<option>Produto não encontrado</option></select>"
		End If
		Response.Write retorno
	End Sub

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
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "174.37.245.58"
			 
			'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			 
			'PORTA DO CDO
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			 
			'TEMPO DE TIMEOUT (EM SEGUNDOS)
			objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
			 
			'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
			objCDOSYSCon.Fields.update
			Set objCDOSYSMail.Configuration = objCDOSYSCon
			 
			'NOME DO REMETENTE, E-MAIL DO REMETENTE
			objCDOSYSMail.From = "Etanji <etanji@okidata.com.br>"
			 
			'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
			objCDOSYSMail.To = Destino
			 
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
	
	Sub AprovarSolColeta()
		Dim sSql
		dim sql, arr, intarr, i
		dim sqltransp
		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
			     "SET [Status_coleta_idStatus_coleta] = 2 " & _
				 ",[data_aprovacao] = GetDate() " & _
				 "WHERE [idSolicitacao_coleta] = " & Request.QueryString("id")
		Call exec(sSql)
		sql = "select " & _
					"c.iscoletaemail, " & _ 
					"c.idtransportadoras " & _
					"from solicitacao_coleta_has_clientes as a " & _
					"left join clientes as b " & _
					"on a.clientes_idclientes  = b.idclientes " & _
					"left join transportadoras as c " & _
					"on b.transportadoras_idtransportadoras = c.idtransportadoras " & _
					"where a.typecolect = 1 and a.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("id")
				
		sql = "select c.iscoletaemail,  c.idtransportadoras " & _ 
					"from solicitacao_coleta_has_pontos_coleta as a " & _ 
					"left join pontos_coleta as b on a.pontos_coleta_idpontos_coleta  = b.idPontos_coleta" & _ 
					"left join transportadoras as c on b.idtransp = c.idtransportadoras " & _ 
					"where a.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("id")

		call search(sql, arr, intarr)
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
		Response.Write "Solicitação aprovada com sucesso!"				 
	End Sub

	Sub AprovarSolColetaCli()
		Dim sSql
		dim sql, arr, intarr, i
		dim sqltransp
		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
			     "SET [Status_coleta_idStatus_coleta] = 2 " & _
				 ",[data_aprovacao] = GetDate() " & _
				 "WHERE [idSolicitacao_coleta] = " & Request.QueryString("id")
		Call exec(sSql)
		sql = "select " & _
					"c.iscoletaemail, " & _ 
					"c.idtransportadoras " & _
					"from solicitacao_coleta_has_clientes as a " & _
					"left join clientes as b " & _
					"on a.clientes_idclientes  = b.idclientes " & _
					"left join transportadoras as c " & _
					"on b.transportadoras_idtransportadoras = c.idtransportadoras " & _
					"where a.typecolect = 1 and a.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("id")
				
		call search(sql, arr, intarr)
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
		Response.Write "Solicitação aprovada com sucesso!"				 
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
