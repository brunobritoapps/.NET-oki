<%
	'|--------------------------------------------------------------------
	'| Arquivo: db.asp																									 
	'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
	'| Data Criação: 13/04/2007																					 
	'| Data Modificação : 15/04/2007																		 
	'| Descrição: Arquivo que Administra a Conexão com o Banco de Dados	 
	'|--------------------------------------------------------------------
	
	'//--------------------------------------------1----------------------------------------------------------------------------------
	'// Variaveis Globais do Arquivo
	'//------------------------------------------------------------------------------------------------------------------------------
	Dim oConn
	Dim oEmail
	Dim ipServer
	Dim userDB
	Dim passwordDB
	Dim serverDB
	Dim databaseDB
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Servidor onde se encontra Hospedado
	'//------------------------------------------------------------------------------------------------------------------------------
	ipServer = Request.ServerVariables("REMOTE_ADDR")
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Set de Objetos Utilizados
	'//------------------------------------------------------------------------------------------------------------------------------
	Set oConn = Server.CreateObject("ADODB.Connection")
'	Set oEmail = Server.CreateObject("CDONTS.NewMail")
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Configurações de Email do Log
	'//------------------------------------------------------------------------------------------------------------------------------
'	oEmail.To = "leandro@miracula.com.br"
'	oEmail.From = "leandro@miracula.com.br"
'	oEmail.Subject = "Log Okidata - SGRS"
'	oEmail.Importance = 1
'	oEmail.BodyFormat = 0
'	oEmail.MailFormat = 0
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Configurações da Camada de Dados
	'//------------------------------------------------------------------------------------------------------------------------------
	
	'Response.Write Left(ipServer, 3) & "<hr>"
	'Response.End
	
	If Left(ipServer, 3) = "127" Then
		userDB = "lusia"
		passwordDB = "123456"
		databaseDB = "marketingoki2"
		serverDB = "MICRO"
	ElseIf Left(ipServer, 3) = "192" Then 	
		userDB = "sa"
		passwordDB = "wealucas"
		databaseDB = "marketingoki2"
		serverDB = "quata"
	Else
		userDB = "Nexxia"
		passwordDB = "nxi4aex7"
		databaseDB = "marketingoki2"
		serverDB = "NEXXIA"
	End If

'		oConn.ConnectionString = "Provider=SQLOLEDB.1;User ID="&userDB&";" & _
'														 "Password="&passwordDB&";" & _
'														 "Persist Security Info=True;" & _
'														 "Initial Catalog="&databaseDB&";" & _
'														 "Data Source="&serverDB
		
		oConn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=lusia;Initial Catalog=marketingoki2;Data Source=MICRO"
		
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Procedimento : Abre Conexão
	'//------------------------------------------------------------------------------------------------------------------------------
	Sub open()
		On Error Resume Next

		If oConn.State <> 1 Then
			oConn.Open
		End If
		
		If Err <> 0 Then
			Response.Write Err.Description
			Response.End()
'			Call log(Err.Description)				
		End If
	End Sub
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Procedimento : Fecha Conexão
	'//------------------------------------------------------------------------------------------------------------------------------
	Sub close()
		On Error Resume Next
		
		If oConn.State <> 0 Then
			oConn.Close		
		End If
		
		If Err <> 0 Then
			Response.Write Err.Description
			Response.End()
'			Call log(Err.Description)				
		End If
	End Sub
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Procedimento : Faz Consulta
	'// Parâmetros :-
	'//  - sSql = Query a ser consultada
	'//  - arrValores = Array que acumulará os registros
	'//  - intValores = Inteiro que guardará quantos registros a consulta possui 
	'//------------------------------------------------------------------------------------------------------------------------------
	Sub search(sSql, arrValores, intValores)		
		On Error Resume Next
		open()
		Dim oRecord

		Set oRecord = Server.CreateObject("ADODB.Recordset")
		
		oRecord.ActiveConnection = oConn
		oRecord.Open sSql

		'response.write oRecord.EOF & "<hr>"
		'response.end		

		If Not oRecord.EOF Then
			arrValores = oRecord.GetRows
			intValores = Ubound(arrValores, 2)			
		Else
			intValores = -1
		End If

		oRecord.Close

		If Err <> 0 Then
			'oRecord.Close
			Response.Write "Erro na sub SEARCH do arquivo db.asp<p>Nº do erro: "&err.number&"<p>Descrição: "&Err.Description &"<p>Query: "&sSql
			Response.End()
'			Call log(Err.Description)				
		End If
	End Sub
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Procedimento : Executa Query
	'// Parâmetros :-
	'//  - sSql = Query a ser executada
	'//------------------------------------------------------------------------------------------------------------------------------
	Sub exec(sSql)
		On Error Resume Next
		
		oConn.Execute(sSql)
		
		If Err <> 0 Then
			Response.Write Err.Description
			Response.End()
'			Call log(Err.Description)				
		End If
	End Sub
	'//------------------------------------------------------------------------------------------------------------------------------
	'// Procedimento : Envia Email com log
	'// Parâmetros :-
	'//  - sError = Err.Description que tem que ser passado
	'//------------------------------------------------------------------------------------------------------------------------------
'	Sub log(sError)
'		oEmail.Body = sError
'		oEmail.Send
'	End Sub
	'//------------------------------------------------------------------------------------------------------------------------------
'********************************************************
'Nome da função: fncPaginar
'Desenvolvido por: Wellington
'Descrição: Função de paginação, para que o usuário possa fazer paginação dos resultados de consulta
'********************************************************
'atualizamos numero de pagina 
function fPaginar(pSql, pagina)

If Request.QueryString("pag")<>"" Then 
   Session("pagina")=Request.QueryString("pag") 
Else 
   Session("pagina")=1 
End If 

'constantes ADO VBScript 
Const adCmdText = &H0001 
Const adOpenStatic = 3 

'Set Conn = Server.CreateObject("ADODB.Connection") 
Set Command = Server.CreateObject("ADODB.Command") 
Set RS =Server.CreateObject("ADODB.RecordSet") 

call open()
session("sql") = pSql

RS.Open pSql,oConn,adopenstatic,adcmdtext

'resultados por pagina a escolher arbitrariamente 
num_registros = 20

'Dimensionar as paginas e determinar a pagina atual 
if not rs.eof then
	RS.PageSize=num_registros
	RS.AbsolutePage=Session("pagina") 
end if

registros_mostrados = 0 

While (Not RS.eof And registros_mostrados < num_registros) 
   registros_mostrados = registros_mostrados +1 
   RS.MoveNext 
Wend 

i=0

While i<RS.PageCount 
  i=i+1 
	if CINT(i) <> CINT(Request.QueryString("pag"))  then
		sTexto = sTexto & "<a href='"&pagina&"?pag="&i&"' class='linkOperacional'> "&i&" </a>  <fon class='linkOperacional'>|</font> "
	else
		sTexto = sTexto & i&" | "
	end if
Wend

RS.Close 
Close() 

set rs = nothing
set Conn= nothing
fPaginar= sTexto

'***************************************
'Fim paginação
'***************************************
end function
	
	Function Paginacao(lLimite, lTotRegs, lPagAtual, sPagASP, sParamURL)
		Dim limite, pagina, MaxPages, counter, sLinkPag
		Dim vlMaximo, vlMinimo

		limite  = lLimite

		if lTotRegs < limite then exit function

		if (lTotRegs/limite) > 1 and (lTotRegs/limite) < 2 then
			MaxPages = 2
		else
			MaxPages = Round(lTotRegs/limite)
		end if

		pagina = lPagAtual
		
		sParamURL = replace(sParamURL, "pag=" & pagina, "")		
		
		If pagina = "" Then
			pagina = 1
		Else
			pagina = pagina
		End If

		counter = 1
		Do While counter <= MaxPages
			if cint(counter) = cint(pagina) then
				Paginacao = Paginacao & counter
			else				
				if len(trim(sParamURL)) then
					sParamURL = "&" & sParamURL
				end if
				sParamURL = replace(sParamURL, "&&", "&")
				Paginacao = Paginacao & " <a href=" & sPagASP & ".asp?pag=" & counter & sParamURL & " class='linkOperacional'> " & Counter & "</a>"				
			end if
					
			If counter < MaxPages Then
				Paginacao = Paginacao & " | "
			End If					
		  counter = counter + 1
		Loop
		
		'Response.Write Paginacao
	End Function
	
	Function vlrMinimo(lLimite, lTotRegs, lPagAtual)
		Dim limite, pagina

		limite  = lLimite
		
		if lTotRegs < limite then
			vlrMinimo = 0
			exit function
		end if
		
		pagina = lPagAtual

		IF pagina = "" or pagina = 1 Then
			IF limite < lTotRegs Then
			  vlrMinimo = 0
			Else
			  vlrMinimo = lTotRegs
			End IF
		Else
			vlrMinimo = ((pagina-1)*(limite))+1
		End IF
	End Function

	Function vlrMaximo(lLimite, lTotRegs, lPagAtual)
		Dim limite, pagina

		limite  = lLimite

		if lTotRegs < limite then
			vlrMaximo = lTotRegs
			exit function
		end if

		pagina = lPagAtual

		If pagina = "" or pagina = 1 Then
			If limite < lTotRegs Then
			  vlrMaximo = limite
			Else
			  vlrMaximo = lTotRegs
			End If
		Else
			vlrMaximo = (limite)*(pagina)
			'vlrMaximo = ((limite-1)*(pagina+1))+((pagina)*1)			
			If vlrMaximo > lTotRegs Then vlrMaximo = lTotRegs
		End If		
	End Function	
	
  Function PaginacaoExibir(intPagina, numeroPorPagina, numItens)
	  Dim mstrRet, _
		     mintA, _
			 mstrRef, _
	         mintPgIni, _
			 mintPgFim, _
			 nOffset, _
			 nIndice
				
		Dim l_intPaginas, _
		    l_intTotal, _
				l_intItemIni, _
				l_intItemFim
				
		Dim paginasPorVez
		paginasPorVez = 10

		If (intPagina * numeroPorPagina) > numItens Then
			l_intTotal = numItens - ((intPagina - 1) * numeroPorPagina)
			l_intItemFim = numItens
		Else
			l_intPaginas = numeroPorPagina
			l_intItemFim = intPagina * numeroPorPagina
		End If
		
		l_intItemIni = (intPagina - 1) * numeroPorPagina + 1
		l_intPaginas = ((numItens - (numItens mod numeroPorPagina)) / numeroPorPagina)' + 1
		
		if (numItens mod numeroPorPagina) > 0 Then
			l_intPaginas = l_intPaginas + 1 
		End if
		
		'Sai caso não tenha páginas suficientes
		If l_intPaginas < 2 Then
			PaginacaoExibir = ""
			Exit Function
		End If


	  'Obtém a página pra onde deve apontar
	  mstrRet = Space(255)
	  mstrRef = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString
	  mstrRef = Mid(mstrRef, 1, XIF(InStr(mstrRef, "pg=") > 0, InStr(mstrRef, "pg=") - 1, Len(mstrRef)))
		
		if right(mstrRef,1) <> "?" and right(mstrRef,1) <> "&" then
			mstrRef = mstrRef & "&"
		end if
		'Response.Write mstrRef & "<hr>"

	 'Obtendo número da páginas possíveis
	  nOffset        = CInt(paginasPorVez / 2)

	  If intPagina + nOffset > l_intPaginas Then
	    nOffset = l_intPaginas - intPagina
	    mintPgIni = (l_intPaginas - numeroPorPagina) + 1
	  Else
	    mintPgIni = intPagina - nOffset
	  End If
		
	  If mintPgIni <= 0 Then mintPgIni = 1
	  mintPgFim = intPagina + nOffset
		
	  'Cria string com paginação
	  PaginacaoExibir =                  "                 <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                    <tr>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                      <td><font size=""1"" face=""Verdana, Arial, Helvetica, sans-serif"">P&aacute;gina " & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                       <strong> <font color=""C00000"">" & intPagina & "</font></strong> de " & l_intPaginas & "<br>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                       Exibindo " & XIF(l_intPaginas = 1, "ítem", "ítens") & " de <strong>" & l_intItemIni & "</strong>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                        a <strong>" & l_intItemFim & "</strong>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                        de um total de " & numItens & "<br>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                        </font>" & vbcrlf
		
	  PaginacaoExibir = PaginacaoExibir & "                         <table border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/fundo_paginacao.gif"">" & vbcrlf
	  PaginacaoExibir = PaginacaoExibir & "                           <tr>" & vbcrlf
	  PaginacaoExibir = PaginacaoExibir & "                             <td width=""25"">" & vbcrlf
		

	  If CInt(intPagina) > 1 Then
			PaginacaoExibir = PaginacaoExibir & "<a href=""" & mstrRef & "pg=" & intPagina - 1 & """><img src=""img/btn_paginacao_anterior.gif"" border=""0"" align=""top""></a>"
			PaginacaoExibir = PaginacaoExibir & "</td>" & vbcrlf
			PaginacaoExibir = PaginacaoExibir & "                            <td width=""5""><img src=""img/paginacao_cantoE.gif"" width=""5"" height=""5""></td>" & vbcrlf

		Else

		    'PaginacaoExibir = PaginacaoExibir & "<img src=""img/btn_paginacao_anterior.gif"" >"
			PaginacaoExibir = PaginacaoExibir & "</td>" & vbcrlf
			PaginacaoExibir = PaginacaoExibir & "                            <td width=""5""></td>" & vbcrlf
			'<img src=""img/paginacao_cantoEOn.gif"" width=""5"" height=""5"">

	  End If
		
	  For nIndice = mintPgIni To mintPgFim
		  PaginacaoExibir = PaginacaoExibir & "                            <td onmouseover=""this.style.cursor='hand'"" onclick=""window.location.href='" & mstrRef & "pg=" & nIndice & "'"">" & aux_DataEscrever(nIndice,CBool(CInt(nIndice) = CInt(intPagina)),(nIndice = l_intPaginas)) & "</td>" & vbcrlf
	  Next

		

	  'Aqui vê se deve ativar ou não o botão "próxima"
	  If CInt(intPagina) < CInt(l_intPaginas) Then
	    PaginacaoExibir = PaginacaoExibir & "                            <td width=""5""><img src=""img/paginacao_cantoD.gif"" width=""5"" height=""5""></td>" & vbcrlf
	    PaginacaoExibir = PaginacaoExibir & "                            <td width=""76"">&nbsp;&nbsp;&nbsp;&nbsp;<a href=""" & mstrRef & "pg=" & intPagina + 1 & """><img src=""img/btn_paginacao_proxima.gif"" border=""0"" align=""top""></a></td>" & vbcrlf
		Else
	    PaginacaoExibir = PaginacaoExibir & "                            <td width=""5""></td>" & vbcrlf
	'    <img src=""img/paginacao_cantoDOn.gif"" width=""5"" height=""17"">
		
		PaginacaoExibir = PaginacaoExibir & "                            <td width=""76""></td>" & vbcrlf
	'	<img src=""img/btn_paginacao_proxima.gif"" border=""0"">
	  End If

	  PaginacaoExibir = PaginacaoExibir & "                          </tr>" & vbcrlf
	  PaginacaoExibir = PaginacaoExibir & "                        </table>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                    </td></tr>" & vbcrlf
		PaginacaoExibir = PaginacaoExibir & "                  </table>" & vbcrlf
	End Function	
	
	Function XIF(par_blnCondicao, par_strVerdadeiro, par_strFalso)
	  Dim l_strRetorno
	  
		If IsNull(par_blnCondicao) Then
		  l_strRetorno = par_strFalso
		Else
			If CBool(par_blnCondicao) Then
				l_strRetorno = par_strVerdadeiro
			Else
				l_strRetorno = par_strFalso
			End If
		End If
	  
	  XIF = l_strRetorno
	End Function
	
	Function aux_DataEscrever(intPagina, blnAtual, blnUltima)
	  Dim blnPrimeira
	  blnPrimeira = (intPagina = 1)
		
	  If blnAtual Then
			aux_DataEscrever =                   "                              <table width=""18"" height=""17"" border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/fundo_paginacaoOn.gif"">" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                <tr>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""1"" bgcolor=""#ffffff""><img src=""img/_spacer.gif"" width=""1"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""2""><img src=""img/_spacer.gif"" width=""2"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td align=""center""><font size=""1"" face=""Arial, Helvetica, sans-serif"">" & intPagina & "</font></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""2""><img src=""img/_spacer.gif"" width=""2"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""1"" bgcolor=""C56C66""><img src=""img/_spacer.gif"" width=""1"" height=""1""></td>" & vbcrlf
		Else
			aux_DataEscrever =                   "                              <table width=""18"" height=""17"" border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/fundo_paginacao.gif"">" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                <tr>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""1"" bgcolor=""#ffffff""><img src=""img/_spacer.gif"" width=""1"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""2""><img src=""img/_spacer.gif"" width=""2"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td align=""center""><font size=""1"" face=""Arial, Helvetica, sans-serif"">" & intPagina & "</font></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""2""><img src=""img/_spacer.gif"" width=""2"" height=""1""></td>" & vbcrlf
			aux_DataEscrever = aux_DataEscrever & "                                  <td width=""1"" bgcolor=""919191""><img src=""img/_spacer.gif"" width=""1"" height=""1""></td>" & vbcrlf
		End If
		
		aux_DataEscrever = aux_DataEscrever & "                                </tr>" & vbcrlf
	  aux_DataEscrever = aux_DataEscrever & "                              </table>" & vbcrlf
	End Function

	Function PaginacaoExibirSaldo(intPagina, numeroPorPagina, numItens, sUrl)
	  Dim mstrRet, _
		     mintA, _
			 mstrRef, _
	         mintPgIni, _
			 mintPgFim, _
			 nOffset, _
			 nIndice
				
		Dim l_intPaginas, _
		    l_intTotal, _
				l_intItemIni, _
				l_intItemFim
				
		Dim paginasPorVez
		paginasPorVez = 10

		If (intPagina * numeroPorPagina) > numItens Then
			l_intTotal = numItens - ((intPagina - 1) * numeroPorPagina)
			l_intItemFim = numItens
		Else
			l_intPaginas = numeroPorPagina
			l_intItemFim = intPagina * numeroPorPagina
		End If
		
		l_intItemIni = (intPagina - 1) * numeroPorPagina + 1
		l_intPaginas = ((numItens - (numItens mod numeroPorPagina)) / numeroPorPagina)' + 1
		
		if (numItens mod numeroPorPagina) > 0 Then
			l_intPaginas = l_intPaginas + 1 
		End if
		
		'Sai caso não tenha páginas suficientes
		If l_intPaginas < 2 Then
			PaginacaoExibir = ""
			Exit Function
		End If


	  'Obtém a página pra onde deve apontar
	  mstrRet = Space(255)
	  mstrRef = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString
	  mstrRef = Mid(mstrRef, 1, XIF(InStr(mstrRef, "pg=") > 0, InStr(mstrRef, "pg=") - 1, Len(mstrRef)))
		
		if right(mstrRef,1) <> "?" and right(mstrRef,1) <> "&" then
			mstrRef = mstrRef & "&"
		end if
		'Response.Write mstrRef & "<hr>"

	 'Obtendo número da páginas possíveis
	  nOffset        = CInt(paginasPorVez / 2)

	  If intPagina + nOffset > l_intPaginas Then
	    nOffset = l_intPaginas - intPagina
	    mintPgIni = (l_intPaginas - numeroPorPagina) + 1
	  Else
	    mintPgIni = intPagina - nOffset
	  End If
		
	  If mintPgIni <= 0 Then mintPgIni = 1
	  mintPgFim = intPagina + nOffset
		
	  'Cria string com paginação
	  PaginacaoExibirSaldo =                  "                 <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                    <tr>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                      <td><font size=""1"" face=""Verdana, Arial, Helvetica, sans-serif"">P&aacute;gina " & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                       <strong> <font color=""C00000"">" & intPagina & "</font></strong> de " & l_intPaginas & "<br>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                       Exibindo " & XIF(l_intPaginas = 1, "ítem", "ítens") & " de <strong>" & l_intItemIni & "</strong>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                        a <strong>" & l_intItemFim & "</strong>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                        de um total de " & numItens & "<br>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                        </font>" & vbcrlf
		
	  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                         <table border=""0"" cellpadding=""0"" cellspacing=""0"" background=""img/fundo_paginacao.gif"">" & vbcrlf
	  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                           <tr>" & vbcrlf
	  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                             <td width=""25"">" & vbcrlf
		

	  If CInt(intPagina) > 1 Then
			PaginacaoExibirSaldo = PaginacaoExibirSaldo & "<a href=""" & mstrRef & "pg=" & intPagina - 1 & sUrl & """><img src=""img/btn_paginacao_anterior.gif"" border=""0"" align=""top""></a>"
			PaginacaoExibirSaldo = PaginacaoExibirSaldo & "</td>" & vbcrlf
			PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""5""><img src=""img/paginacao_cantoE.gif"" width=""5"" height=""5""></td>" & vbcrlf

		Else

		    'PaginacaoExibirSaldo = PaginacaoExibirSaldo & "<img src=""img/btn_paginacao_anterior.gif"" >"
			PaginacaoExibirSaldo = PaginacaoExibirSaldo & "</td>" & vbcrlf
			PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""5""></td>" & vbcrlf
			'<img src=""img/paginacao_cantoEOn.gif"" width=""5"" height=""5"">

	  End If
		
	  For nIndice = mintPgIni To mintPgFim
		  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td onmouseover=""this.style.cursor='hand'"" onclick=""window.location.href='" & mstrRef & "pg=" & nIndice & sUrl &"'"">" & aux_DataEscrever(nIndice,CBool(CInt(nIndice) = CInt(intPagina)),(nIndice = l_intPaginas)) & "</td>" & vbcrlf
	  Next

		

	  'Aqui vê se deve ativar ou não o botão "próxima"
	  If CInt(intPagina) < CInt(l_intPaginas) Then
	    PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""5""><img src=""img/paginacao_cantoD.gif"" width=""5"" height=""5""></td>" & vbcrlf
	    PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""76"">&nbsp;&nbsp;&nbsp;&nbsp;<a href=""" & mstrRef & "pg=" & intPagina + 1 & sUrl &"""><img src=""img/btn_paginacao_proxima.gif"" border=""0"" align=""top""></a></td>" & vbcrlf
		Else
	    PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""5""></td>" & vbcrlf
	'    <img src=""img/paginacao_cantoDOn.gif"" width=""5"" height=""17"">
		
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                            <td width=""76""></td>" & vbcrlf
	'	<img src=""img/btn_paginacao_proxima.gif"" border=""0"">
	  End If

	  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                          </tr>" & vbcrlf
	  PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                        </table>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                    </td></tr>" & vbcrlf
		PaginacaoExibirSaldo = PaginacaoExibirSaldo & "                  </table>" & vbcrlf
	End Function
%>
