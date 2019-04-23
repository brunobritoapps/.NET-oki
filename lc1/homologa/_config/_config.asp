<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="db.asp" -->
<%
	'|--------------------------------------------------------------------
	'| Arquivo: _config.asp
	'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)
	'| Data Criação: 13/04/2007
	'| Data Modificação : 15/04/2007
	'| Descrição: Arquivo de Administra as Configurações e contém _
	'| todos os includes de Configuração
	'|--------------------------------------------------------------------

	Dim URL
	Dim TITLE

	Session.Timeout = 20

	If Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "127" Then
		TITLE = "Sistema de Gerenciamento de Retorno de Suprimentos Vazios(Local)"
		URL = "http://localhost:81/sgrs/"
	ElseIf Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "192" Then
		TITLE = "Sistema de Gerenciamento de Retorno de Suprimentos Vazios(Local)"
		URL = "http://192.168.0.1:85/"
	Else
		TITLE = "Sistema de Gerenciamento de Retorno de Suprimentos Vazios"
		URL = "http://www.sustentabilidadeoki.com.br/"
		'URL = "http://ftpodb.okidata.com.br/sustentabilidade/"
	End If

	'============================================================================================
	'| Function que valida o Digito Verificador da Solicitação de Coleta
	'============================================================================================
	Function getDigitoControle(NumeroSolicitacaoColeta)
		Dim arrCont(9)
		Dim arrCont2(10)
		Dim i, j
		Dim sumFirstDigit
		Dim sumSecondDigit
		Dim numeroSolicitacao
		Dim firstDig
		Dim secondDig
		numeroSolicitacao = Mid(NumeroSolicitacaoColeta, 2)
		sumFirstDigit = 0
		sumSecondDigit = 0
		'*****************
		arrCont(0) = 2
		arrCont(1) = 9
		arrCont(2) = 8
		arrCont(3) = 7
		arrCont(4) = 6
		arrCont(5) = 5
		arrCont(6) = 4
		arrCont(7) = 3
		arrCont(8) = 2
		'*****************
		arrCont2(0) = 3
		arrCont2(1) = 2
		arrCont2(2) = 9
		arrCont2(3) = 8
		arrCont2(4) = 7
		arrCont2(5) = 6
		arrCont2(6) = 5
		arrCont2(7) = 4
		arrCont2(8) = 3
		arrCont2(9) = 2
		'*****************
		For i=1 To UBound(arrCont)
			sumFirstDigit = sumFirstDigit + (Mid(numeroSolicitacao, i, 1) * arrCont(i - 1))
		Next
		If (sumFirstDigit Mod 11) < 2 Then
			firstDig = 0
		Else
			firstDig = 11 - (sumFirstDigit Mod 11)
		End If
		numeroSolicitacao = numeroSolicitacao & firstDig
		For j=1 To UBound(arrCont2)
			sumSecondDigit = sumSecondDigit + (Mid(numeroSolicitacao, j, 1) * arrCont2(j - 1))
		Next
		If (sumSecondDigit Mod 11) < 2 Then
			secondDig = 0
		Else
			secondDig = 11 - (sumSecondDigit Mod 11)
		End If
		NumeroSolicitacaoColeta = Left(NumeroSolicitacaoColeta, 1) & numeroSolicitacao & secondDig
		getDigitoControle = NumeroSolicitacaoColeta
	End Function
	'============================================================================================
	'| Function que gera o número Sequencial que terá que ser atualizado pelo cliente
	'============================================================================================
	Function getSequencial(bFirst)
		Dim numSequencial, dataNumSequencial
		Dim sSql, arrNumSeqIDCliente, intNumSeqIDCliente, i

		sSql = "select count([idSolicitacao_coleta]) from solicitacao_coleta where data_solicitacao between '"&getDateLess()&" 00:00:00' and '"&getDateMore()&" 23:59:59'"
'		response.write sSql
'		response.end
'		sSql = "SELECT COUNT([idSolicitacao_coleta]) FROM [marketingoki2].[dbo].[Solicitacao_coleta]"
		Call search(sSql, arrNumSeqIDCliente, intNumSeqIDCliente)
		If intNumSeqIDCliente > -1 Then
			if arrNumSeqIDCliente(0,i) <> 0 and arrNumSeqIDCliente(0,i) <> "" then
				For i=0 To intNumSeqIDCliente
					numSequencial = IncrementNumber(arrNumSeqIDCliente(0,i), bFirst)
				Next
			else
				numSequencial = "00001"
			end if
		Else
			numSequencial = "00001"
		End If
		getSequencial = numSequencial

	End Function

	Sub getSessionUser()
		if request.QueryString("logoff") then
			session.Abandon()
			Session("IDContato") = ""
			response.redirect "index.asp?area=home"
		end if
		If Session("IDContato") = "" Then
			Response.Redirect "frmLoginCliente.asp?IDMessage=Favor Identifique-se!"
		End If
	End Sub

	Sub UpdateSessions()
		Dim sSql, arrUser, intUser, i
		Dim sSql2, arrPontoCliente, intPontoCliente
		Dim sSql3, arrNumeroSequencial, intNumeroSequencial

		sSql = "SELECT " & _
					 "[A].[idContatos], " & _
					 "[A].[Clientes_idClientes], " & _
					 "[A].[nome], " & _
					 "[A].[email], " & _
					 "[A].[isMaster], " & _
					 "[C].[isColetaDomiciliar] " & _
					 "FROM [marketingoki2].[dbo].[Contatos] AS [A] " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [B] " & _
					 "ON [A].[Clientes_idClientes] = [B].[idClientes] " & _
					 "LEFT JOIN [marketingoki2].[dbo].[Categorias] AS [C] " & _
					 "ON [B].[Categorias_idCategorias] = [C].[idCategorias] " & _
					 "WHERE [A].[status_contato] = 1 " & _
					 "AND [A].[usuario] = '"& Session("User") &"' " & _
					 "AND [A].[senha] = '"& Session("Password") &"'"

		Call search(sSql, arrUser, intUser)
		If intUser > -1 Then
			For i=0 To intUser
				Session("IDContato") = arrUser(0,i)
				Session("IDCliente") = arrUser(1,i)
				Session("NomeContato") = arrUser(2,i)
				Session("Email") = arrUser(3,i)
				Session("isMaster") = arrUser(4,i)
				Session("isColetaDomiciliar") = arrUser(5,i)

				sSql3 = "SELECT numero_sequencial, data_atualizacao_sequencial FROM Clientes WHERE idClientes = " & Session("IDCliente")
				Call search(sSql3, arrNumeroSequencial, intNumeroSequencial)
				If intNumeroSequencial > -1 Then
					Session("NumeroSequencial") = arrNumeroSequencial(0,0)
					Session("DataSequencial") = arrNumeroSequencial(1,0)
				End If

				If Session("isColetaDomiciliar") = 0 Then
					sSql2 = "SELECT * FROM Solicitacao_coleta_has_Clientes WHERE Clientes_idClientes = " & Session("IDCliente")
					Call search(sSql2, arrPontoCliente, intPontoCliente)
					If intPontoCliente < -1 Then
						Session("IDPontoColeta") = arrPontoCliente(1,0)
					End If
				End If
			Next
		End If
	End Sub

	Function IncrementNumber(StringNumber, bFirst)
		Dim DiffNumber
		Dim NewNumber
		Dim i
		Dim sSql

		StringNumber = StringNumber + 1
		DiffNumber = 5 - Len(StringNumber)
		For i=1 To DiffNumber
			NewNumber = NewNumber & "0"
		Next
		NewNumber = NewNumber & StringNumber
		If Not bFirst Then
			If Session("IDCliente") <> "" Then
				sSql = "UPDATE [marketingoki2].[dbo].[Clientes] " & _
						 	 "SET [numero_sequencial] = '" & NewNumber & "' " & _
						 	 "WHERE [idClientes] = " & Session("IDCliente")
				Call exec(sSql)
			End If
		End If

		IncrementNumber = NewNumber
	End Function

	Sub GetSessionAdm()
		If Session("IDAdministrator") = "" Then
			Response.Redirect "frmloginadm.asp?msg=Sessão inválida, por favor efetue a login novamente."
		End If
	End Sub

	Sub GetSessionPonto()
		if request.QueryString("logoff") then
			session.Abandon()
			Session("IDPonto") = ""
			response.redirect "frmloginpc.asp"
		end if
		If Session("IDPonto") = "" Then
			Response.Redirect "frmloginpc.asp?msg=Sessão inválida, por favor efetue a login novamente."
		End If
	End Sub

	function getDateMore()
		dim dia
		dim mes
		dim ano

		mes = month(now())
		ano = year(now())

		select case(mes)
			case 1
				dia = "31"
			case 3
				dia = "31"
			case 5
				dia = "31"
			case 7
				dia = "31"
			case 8
				dia = "31"
			case 10
				dia = "31"
			case 12
				dia = "31"
			case 4
				dia = "30"
			case 6
				dia = "30"
			case 9
				dia = "30"
			case 11
				dia = "30"
			case 2
				dia = "28"
		end select

		if len(mes) = 1 then
			mes = "0" & mes
		end if

		getDateMore = ano & "/" & mes & "/" & dia
	end function

	function getDateLess()
		dim dia
		dim mes
		dim ano

		dia = "01"
		mes = month(now())
		ano = year(now())

		if len(mes) = 1 then
			mes = "0" & mes
		end if

		getDateLess = ano & "/" & mes & "/" & dia
	end function

	function getTextoByArea(area)
		dim sql, arr, intarr, i
		sql = "SELECT [idtexto] " & _
				  ",[texto] " & _
				  ",[area] " & _
			  "FROM [marketingoki2].[dbo].[Home_Textos] where area = '"&area&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getTextoByArea = arr(1,i)
			next
		else
			getTextoByArea = ""
		end if
	end function

	function getNoticiasById(id)
		dim sql, arr, intarr, i
		dim html

		sql = "SELECT [idnoticia] " & _
				  ",[titulo] " & _
				  ",[text] " & _
				  ",[data] " & _
				  ",[fonte] " & _
				  ",[ativo] " & _
			  "FROM [marketingoki2].[dbo].[Home_Noticias] where idnoticia = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				html = html & "<div>"
				html = html & "<div><b>"&arr(1,i)&"</b></div>"
				html = html & arr(2,i)
				html = html & "<div>fonte: <a href="""&arr(4,i)&""">"&arr(4,i)&"</a> / data: "&day(arr(3,i))&"/"&month(arr(3,i))&"/"&year(arr(3,i))&"</div>"
				html = html & "</div>"
			next
		else
			html = ""
		end if
		getNoticiasById = html
	end function

	function getNoticias()
		dim sql, arr, intarr, i
		dim html
		html = ""
		sql = "SELECT [idnoticia] " & _
				  ",[titulo] " & _
				  ",[text] " & _
				  ",[data] " & _
				  ",[fonte] " & _
				  ",[ativo] " & _
			  "FROM [marketingoki2].[dbo].[Home_Noticias] where ativo = 1"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if left(request.servervariables("REMOTE_ADDR"), 3) = "127" then
					html = html & "<a href=""http://localhost:81/sgrs/index.asp?idnoticia="&arr(0,i)&""">"&arr(1,i)&" - "&DateRight(arr(3,i))&"</a><br ><br />"
				elseif left(request.servervariables("REMOTE_ADDR"), 3) = "192" then
					html = html & "<a href=""http://192.168.0.1:85/index.asp?idnoticia="&arr(0,i)&""">"&arr(1,i)&" - "&DateRight(arr(3,i))&"</a><br /><br />"
				else
					html = html & "<a href=""http://www.sustentabilidadeoki.com.br/index.asp?idnoticia="&arr(0,i)&""">"&arr(1,i)&" - "&DateRight(arr(3,i))&"</a><br /><br />"
					'html = html & "<a href=""http://ftpodb.okidata.com.br/sustentabilidade/index.asp?idnoticia="&arr(0,i)&""">"&arr(1,i)&" - "&DateRight(arr(3,i))&"</a><br /><br />"					
				end if
			next
		else
			html = ""
		end if
		getNoticias = html
	end function

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		Dia = Left(sData, 2)
		Dia = Replace(Dia, "/", "")
		If Len(Dia) = 1 Then
			Dia = "0" & Dia
		End If
		If Len(Replace(Left(sData, 2), "/", "")) = 1 Then
			Mes = Mid(sData, 3, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		Else
			Mes = Mid(sData, 4, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		End If
		Ano = Right(sData, 4)
		Ano = Replace(Ano, "/", "")
		If Len(Ano) = 1 Then
			Ano = "0" & Ano
		End If
		DateRight = Mes & "/" & Dia & "/" & Ano
	End Function

%>
