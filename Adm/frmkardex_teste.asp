<!--#include file="../_config/_config.asp" -->
<!-- #include file="../_config/freeaspupload.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Response.Expires = -1
	Server.ScriptTimeout = 600

	Dim uploadsDirVar
	Dim diagnostics
	Dim FileUploaded
	Dim QtdRealProduto
	dim log_error

	dim htmlFileImprimir

	log_error = ""

	If Left(Request.ServerVariables("LOCAL_ADDR"),3) = "127" Or Left(Request.ServerVariables("LOCAL_ADDR"),3) = "192" Then
	  uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Kardex\Domiciliar\"
	Else
	  uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Kardex\Domiciliar\"
	End If

	Function TestEnvironment()
			Dim fso, fileName, testFile, streamTest
			TestEnvironment = ""
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If Not fso.FolderExists(uploadsDirVar) Then
					TestEnvironment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
					Exit Function
			End If
			fileName = uploadsDirVar & "\test.txt"
			On Error Resume Next
			Set testFile = fso.CreateTextFile(fileName, True)
			If Err.Number<>0 Then
					TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
					Exit Function
			End If
			Err.Clear
			testFile.Close
			fso.DeleteFile(fileName)
			If Err.Number<>0 Then
					TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
					Exit Function
			End If
			Err.Clear
			Set streamTest = Server.CreateObject("ADODB.Stream")
			If Err.Number<>0 Then
					TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
					Exit Function
			End If
			Set streamTest = Nothing
	End Function

	Function SaveFiles
			Dim Upload, fileName, fileSize, ks, i, fileKey

			Set Upload = New FreeASPUpload
			Upload.Save(uploadsDirVar)

		' If something fails inside the script, but the exception is handled
		If Err.Number <> 0 Then Exit function

			SaveFiles = ""
			ks = Upload.UploadedFiles.keys
			If (UBound(ks) <> -1) Then
					SaveFiles = "<B>Arquivo Importado:</B> "
					For Each fileKey in Upload.UploadedFiles.keys
							SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
							FileUploaded = Upload.UploadedFiles(fileKey).FileName
					next
			Else
					SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
			End If
	End Function

	Sub Submit()
		If ucase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
				diagnostics = TestEnvironment()
				If diagnostics <> "" Then
						Response.Write diagnostics
				End If
		Else
				Response.Write SaveFiles()
				call lerArquivo()
		End If
	End Sub

'	Function LerArquivo()
'		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
'			On Error Resume Next
'			Dim oFso
'			Dim oFile
'			Dim oStream
'			Dim Linha
'			Dim Ret
'			Dim Arr
'			Dim i
'			Dim Cont
'			Dim NumSolicitacao
'			Dim contTimeWhile
'
'			contTimeWhile = 0
'			NumSolicitacao = ""
'			Set oFso = Server.CreateObject("Scripting.FileSystemObject")
'			Set oFile = oFso.GetFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Kardex\Domiciliar\" & FileUploaded)
'			Set oStream = oFile.OpenAsTextStream(1,False)
'
'			While Not oStream.AtEndOfStream
'				contTimeWhile = contTimeWhile + 1
'				Linha = oStream.ReadLine
'				Arr = Split(Linha,";")
'
'				'Arr(0) = numero da solicitacao coleta
'				'Arr(1) = data recebimento
'				'Arr(2) = cod do produto
'				'Arr(3) = qtd do produto
'
''				response.write berror & "<br />"
'				If ValidateSolicitacao(Arr(0)) Then
'					If ValidateProduct(Arr(2)) Then
''						response.write validaStatusSolicitacao(Arr(0), contTimeWhile) & "<br />"
'						if validaStatusSolicitacao(Arr(0), contTimeWhile) then
'							If contTimeWhile = 1 Then
'								NumSolicitacao = Arr(0)
'							Else
'								If NumSolicitacao <> Arr(0) Then
'									NumSolicitacao = Arr(0)
'									QtdRealProduto = 0
'									berror = false
'								End If
'							End If
'							Ret = Ret & "<tr>"
'							For i=0 To Ubound(Arr)
'								If Cont Mod 2 = 0 Then
'									Ret = Ret &  "<td class='classColorRelPar'>" & Arr(i) & "</td>"
'								Else
'									Ret = Ret &  "<td class='classColorRelImpar'>" & Arr(i) & "</td>"
'								End If
'							Next
'							Ret = Ret & "</tr>"
'							Cont = Cont + 1
'							Call AtualizaSol(Arr(0), Arr(1), Arr(2), Arr(3), contTimeWhile)
'						else
'							berror = true
'							log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de atualiza��o de status n�o pode ser efetuado - Solicita��o: ["&Arr(0)&"]</b><br />"
'						end if
'					Else
'						berror = true
'						log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Produto inexistente / Produto: "&Arr(2)&"</b><br />"
'					End If
'				Else
'					berror = true
'					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Solicita��o inexistente / Solicita��o: "&Arr(0)&"</b><br />"
'				End If
'			Wend
'
'			oStream.Close
'			Set oFso = Nothing
'			QtdRealProduto = 0
'			LerArquivo = Ret
'			If Error <> 0 Then
'				Response.Write "<tr><td colspan=""5"">Erro na opera��o de Atualiza��o</td></tr>"
'				Exit Function
'			End If
'		End If
'	End Function

'	Sub AtualizaSol(NumSol, DataReceb, CodProduto, QtdProduto, cont)
'		Dim sSql, arrSol, intSol, i
'		Dim arrSol2, intSol2, j
'		dim sql, arr, intarr
'		QtdRealProduto = CInt(QtdRealProduto) + CInt(QtdProduto)
'
'		if DataReceb <> "" and not berror then
'			sql = "select * from solicitacao_coleta where data_recebimento = convert(datetime,'"&DataReceb&"') and [numero_solicitacao_coleta] = '" &NumSol& "'"
''			response.write sql
''			response.end
'			call search(sql, arr, intarr)
'			if intarr > -1  and cont = 1then
'				berror = true
'				log_error = log_error & "<b style=""color:#FF0000;"">A Solicita��o "&NumSol&" j� foi operacionada!</b><br />"
'			else
'				if validaDataReceb(DataReceb, NumSol) then
'					sSql = "SELECT [idSolicitacao_coleta] " & _
'									"FROM [marketingoki2].[dbo].[Solicitacao_coleta] " & _
'									"WHERE [numero_solicitacao_coleta] = '" &NumSol& "' " & _
'									"AND MONTH(data_solicitacao) = " & Month(Now())
''					response.write sSql
''					response.end
'					Call search(sSql, arrSol, intSol)
'					If intSol > -1 Then
'						For i=0 To intSol
'							sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
'											"SET " & _
'											"[Status_coleta_idStatus_coleta] = 6, " & _
'											"[qtd_cartuchos_recebidos] = "&QtdRealProduto&", " & _
'											"[data_recebimento] = '"&DataReceb&"' " & _
'											"WHERE [idSolicitacao_coleta] = "&arrSol(0,i)
'							Call exec(sSql)
'							sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
'											"[Produtos_idProdutos], " & _
'											"[Solicitacao_coleta_idSolicitacoes_coleta], " & _
'											"[quantidade]) " & _
'											"VALUES( " & _
'											"'"&CodProduto&"', " & _
'											""&arrSol(0,i)&", " & _
'											""&QtdProduto&")"
'							Call exec(sSql)
'							call addKardex(DataReceb, NumSol, CodProduto, QtdProduto)
'						Next
'					End If
'				else
'					berror = true
'					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Data incorreto / Solicita��o ["&NumSol&"]</b><br />"
'				end if
'			end if
'		else
'			berror = true
'			log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Data incorreto / Solicita��o ["&NumSol&"]</b><br />"
'		end if
'	End Sub

	sub lerArquivo()
		dim fs
		dim fileImport
		dim stream
		dim linha

		linha = ""

		set fs = server.createobject("Scripting.FileSystemObject")
		set fileImport = fs.getFile(request.servervariables("APPL_PHYSICAL_PATH") & "adm\Kardex\Domiciliar\" & FileUploaded)
		set stream = fileImport.openAsTextStream(1, false)

		call criaTabelaTemporaria()
		while not stream.atEndOfStream
			linha = stream.readLine
			call insereTabelaTemporaria(linha)
		wend
		call validaFileImport()
		call htmlFile()
		call dropTabelaTemporaria()

		set stream = nothing
		set fileImport = nothing
		set fs = nothing
	end sub

	sub atualizaSolicitacao(numero_solicitacao, data_recebimento, cod_produto, qtd)
		dim sql, arr, intarr, i
		if isMaster(numero_solicitacao) then
			dim qtd_real
			sql = "SELECT [qtd_cartuchos_recebidos] " & _
				  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] where numero_solicitacao_coleta = '"&numero_solicitacao&"'"
			call search(sql, arr , intarr)
			if intarr > -1 then
				for i=0 to intarr
					if arr(0,i) <> "" then
						qtd_real = qtd_real + cint(arr(0,i))
					else
						qtd_real = 0
					end if
				next
			end if
			qtd_real = qtd_real + qtd
			sql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
							"SET " & _
							"[Status_coleta_idStatus_coleta] = 6, " & _
							"[qtd_cartuchos_recebidos] = "&qtd_real&", " & _
							"[data_recebimento] = convert(datetime, '"&formataDataPonto(data_recebimento)&"') " & _
							"WHERE [idSolicitacao_coleta] = "&getIDSolicitacao(numero_solicitacao)
'			response.end
			call atualizaSolicitacaoPonto(numero_solicitacao, cod_produto, qtd, qtd_real, data_recebimento)
'			response.write sql
'			response.end
			call exec(sql)
			sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
							"[Produtos_idProdutos], " & _
							"[Solicitacao_coleta_idSolicitacoes_coleta], " & _
							"[quantidade]) " & _
							"VALUES( " & _
							"'"&cod_produto&"', " & _
							""&getIDSolicitacao(numero_solicitacao)&", " & _
							""&qtd&")"
	'		response.write sql & "<br>"
'			response.write sql
'			response.end
			call exec(sql)
		else
			dim arr2, intarr2, j
			dim qtd_real2
			qtd_real2 = 0
			sql = "SELECT [qtd_cartuchos_recebidos] " & _
				  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] where numero_solicitacao_coleta = '"&numero_solicitacao&"'"
			call search(sql, arr2, intarr2)
			if intarr2 > -1 then
				for j=0 to intarr2
					if arr2(0,j) <> "" then
						qtd_real2 = clng(arr2(0,j))
					else
						qtd_real2 = 0
					end if
				next
			end if
			qtd_real2 = qtd_real2 + clng(qtd)
			sql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
							"SET " & _
							"[Status_coleta_idStatus_coleta] = 6, " & _
							"[qtd_cartuchos_recebidos] = "&qtd_real2&", " & _
							"[data_recebimento] = '"&formatdate(data_recebimento)&"'" & _
							"WHERE [numero_solicitacao_coleta] = '"&numero_solicitacao&"'"
			'response.write sql
			'response.write data_recebimento
			'response.end
			call exec(sql)
			sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
							"[Produtos_idProdutos], " & _
							"[Solicitacao_coleta_idSolicitacoes_coleta], " & _
							"[quantidade]) " & _
							"VALUES( " & _
							"'"&cod_produto&"', " & _
							""&getIDSolicitacao(numero_solicitacao)&", " & _
							""&qtd&")"
'			response.write sql & "<br>"
'			response.write sql
'			response.end
			call exec(sql)
		end if
	end sub

	function formataDataPonto(sdata)
		dim data
		dim dia
		dim mes
		dim ano
		data = split(sdata, "/")
		dia = data(0)
		mes = data(1)
		ano = data(2)
		if len(dia) = 1 then
			dia = "0"&dia
		end if
		if len(mes) = 1 then
			mes = "0"&mes
		end if
		if len(ano) = 1 then
			ano = "0"&ano
		end if
		'formataDataPonto = ano&"/"&mes&"/"&dia
		'formataDataPonto = dia&"/"&mes&"/"&ano  
		formataDataPonto = ano&"-"&mes&"-"&dia  
	end function

	sub atualizaSolicitacaoPonto(numero_solicitacao, cod_produto, qtd, qtd_real, data_recebimento)
		dim sql, arr, intarr, i
		sql = "SELECT [id_solicitacao] " & _
					  ",[id_pontocoleta] " & _
					  ",[numero_solicitacao_master] " & _
					  ",[is_baixada] " & _
				  "FROM [marketingoki2].[dbo].[Solicitacoes_Baixadas] " & _
				  "WHERE [numero_solicitacao_master] = '"&numero_solicitacao&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
'				sql = "DELETE FROM [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos] WHERE [Solicitacao_coleta_idSolicitacoes_coleta] = " & arr(0,i)
'				call exec(sql)
				sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
								"[Produtos_idProdutos], " & _
								"[Solicitacao_coleta_idSolicitacoes_coleta], " & _
								"[quantidade]) " & _
								"VALUES( " & _
								"'"&cod_produto&"', " & _
								""&arr(0,i)&", " & _
								""&qtd&")"
				call exec(sql)
				if cint(arr(3,i)) = 1 then
					sql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
							   "SET [Status_coleta_idStatus_coleta] = 6 " & _
							   ",[qtd_cartuchos_recebidos] = " & qtd_real & " " & _
							   ",[data_recebimento] = convert(datetime, '"&formataDataPonto(data_recebimento)&"') " & _
							 "WHERE [idsolicitacao_coleta] = " & arr(0,i)
'					response.write sql
'					response.end
'			response.write sql
'			response.end
					call exec(sql)
				end if
			next
		end if
	end sub

	sub htmlFile()
		dim sql, arr, intarr, i
		dim style
		htmlFileImprimir = ""
		sql = "select * from ##valida_file_import"
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				htmlFileImprimir = htmlFileImprimir & "<tr>"
				htmlFileImprimir = htmlFileImprimir & "<td "&style&">"&arr(0,i)&"</td>"
				if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
					htmlFileImprimir = htmlFileImprimir & "<td "&style&">"&DateRight(formatdatetime(arr(1,i),2))&"</td>"
				else
					htmlFileImprimir = htmlFileImprimir & "<td "&style&">"&formatdatetime(arr(1,i), 2)&"</td>"
				end if
				htmlFileImprimir = htmlFileImprimir & "<td "&style&">"&arr(2,i)&"</td>"
				htmlFileImprimir = htmlFileImprimir & "<td "&style&">"&arr(3,i)&"</td>"
				htmlFileImprimir = htmlFileImprimir & "</tr>"
			next
		else
			htmlFileImprimir = htmlFileImprimir & "<tr>"
			htmlFileImprimir = htmlFileImprimir & "<td colspan=""4"" "&style&"><b>Nenhum registro encontrado</b></td>"
			htmlFileImprimir = htmlFileImprimir & "</tr>"
		end if
	end sub

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

	sub validaFileImport()
		dim sql, arr, intarr, i
		dim arrValida, intValida, j
		dim numero_solicitacao
		dim berror
		dim cont_acertos

		sql = "select distinct(num_sol_col) from ##valida_file_import"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				sql = "select * from ##valida_file_import where num_sol_col = '"&arr(0,i)&"'"
				numero_solicitacao = arr(0,i)
				cont_acertos = 0
				call search(sql, arrValida, intValida)
				if intValida > -1 then
					for j=0 to intValida
'						response.write "acertos: " & cont_acertos & " limite: " & (intValida + 1) & "<br />"
'						response.write "Solicitacao: " & arrValida(0,j) & " em: " & j & "<br />"
						if j=0 then
							berror = false
							if validaStatusSolicitacao(arrValida(0,j)) then
'								response.write "passou no status <br />"
								if validaRegistro(arrValida(0,j),arrValida(1,j),arrValida(2,j),arrValida(3,j)) then
'									response.write "passou no registro <br />"
									cont_acertos = cont_acertos + 1
								else
									berror = true
								end if
							else
								berror = true
							end if
						else
							if not berror then
'								response.write "passow no berror <br />"
								if validaRegistro(arrValida(0,j),arrValida(1,j),arrValida(2,j),arrValida(3,j)) then
'									response.write "passou no registro 2 <br />"
									cont_acertos = cont_acertos + 1
								end if
							end if
						end if
					next
					if cont_acertos = (intValida + 1) then
						call validaFileImportInsert(numero_solicitacao)
					end if
				end if
			next
		else
			log_error = log_error & "<b style=""color:#FF0000;"">Erro: N�o h� nenhum registro a ser validado</b><br />"
		end if
	end sub

	sub validaFileImportInsert(numsol)
		dim sql, arrValida, intValida, j
		sql = "select * from ##valida_file_import where num_sol_col = '"&numsol&"'"
'		response.write sql & "<br />"
'		response.end
		call search(sql, arrValida, intValida)
		if intValida > -1 then
			for j=0 to intValida
				call addKardex(arrValida(0,j),arrValida(1,j),arrValida(2,j),arrValida(3,j))
			next
			call atualizaProdutoFaltante(numsol)
		end if
	end sub

	sub atualizaProdutoFaltante(numsol)
		dim sql, arr, intarr, i
		dim qtd_recebidos, qtd_cartuchos, data_recebimento, qtd_diferenca
		sql = "select qtd_cartuchos, qtd_cartuchos_recebidos, data_recebimento from solicitacao_coleta where numero_solicitacao_coleta = '"&numsol&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				qtd_cartuchos = clng(qtd_cartuchos) + arr(0,i)
				qtd_recebidos = clng(qtd_recebidos) + arr(1,i)
				data_recebimento = arr(2,i)
			next
		end if
		data_recebimento = formatdate(data_recebimento)
		if qtd_recebidos <> "" and qtd_cartuchos <> "" then
			if clng(qtd_recebidos) < clng(qtd_cartuchos) then
				qtd_diferenca = qtd_cartuchos - qtd_recebidos
'				response.write qtd_diferenca & "<br />"
'				response.write qtd_cartuchos & "<br />"
'				response.write qtd_recebidos & "<br />"
				call addKardex(numsol,data_recebimento,"BR1000001",qtd_diferenca)
				sql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"SET " & _
						"[qtd_cartuchos_recebidos] = "&qtd_recebidos+qtd_diferenca&" " & _
						"WHERE [idSolicitacao_coleta] = "&getIDSolicitacao(numsol)
				call exec(sql)
			end if
		end if
	end sub

	function validaRegistro(num_sol, data_receb, cod_prod, qtd_prod)
		if ValidateSolicitacao(num_sol) then
			'Response.Write month(data_receb) & "/" & day(data_receb) & "/" & year(data_receb) & "<hr>"
			'Response.End
			if validaDataReceb(data_receb, num_sol) then
				if ValidateProduct(cod_prod) then
					validaRegistro = true
				else
					log_error = log_error & "<b style=""color:#FF0000;"">Erro: N�o exite produto com este c�digo ["&cod_prod&"] / N�mero da Solicita��o ["&num_sol&"]</b><br />"
					validaRegistro = false
				end if
			else
				log_error = log_error & "<b style=""color:#FF0000;"">Erro: Data de Recebimento menor que a Data de Envio para Transportadora / N�mero da Solicita��o ["&num_sol&"]</b><br />"
				validaRegistro = false
			end if
		else
			log_error = log_error & "<b style=""color:#FF0000;"">Erro: Solicita��o de Coleta n�mero ["&num_sol&"] n�o existe / N�mero da Solicita��o ["&num_sol&"]</b><br />"
			validaRegistro = false
		end if
	end function

	function montaData(data)
		dim dia
		dim mes
		dim ano

		dia = left(data,2)
		mes = mid(data,3,2)
		ano = right(data,4)
		
		montaData = dia&"/"&mes&"/"&ano
	end function

	sub insereTabelaTemporaria(linha)
		dim sql
		dim arr
		dim x
		
		arr = split(linha, ";")
		
		for each item in Arr
			x = x + 1
		next
		
		if x=4 then
			if linha <> "" then
				if ValidateSolicitacao(arr(0)) then
					sql = "insert into ##valida_file_import values ('"&arr(0)&"', '"&montaData(arr(1))&"', '"&arr(2)&"', "&arr(3)&")"
					call exec(sql)
				else
					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Solicita��o de Coleta n�mero ["&arr(0)&"] n�o existe</b><br />"
				end if
	'			response.write sql & "<br />"
	'			response.end
			end if
		else
			log_error = log_error & "<b style=""color:#FF0000;"">Erro: O Arquivo est� incompleto</b><br />"
		end if
	end sub

	sub criaTabelaTemporaria()
		dim sql
		sql = "create table ##valida_file_import ( " & _
					"num_sol_col varchar(13) not null, " & _
					"data_receb varchar(50) not null, " & _
					"cod_prod varchar(50) not null, " & _
					"qtd_prod int not null)"
		call exec(sql)
	end sub

	sub dropTabelaTemporaria()
		dim sql
		sql = "drop table ##valida_file_import"
		call exec(sql)
	end sub

	Function ValidateProduct(ID)
		Dim sSql, arrProd, intProd, i
		sSql = "select * from produtos where idoki = '"&ID&"'"
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			ValidateProduct = True
		Else
			ValidateProduct = False
		End If
	End Function

	Function ValidateSolicitacao(ID)
		Dim sSql, arrSol, intSol, i
		'sSql = "select * from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"' and MONTH(data_solicitacao) = " & Month(Now())
		sSql = "select * from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"' and status_coleta_idstatus_coleta = 7"
		'response.write sSql
		'response.end
		Call search(sSql, arrSol, intSol)
		If intSol > -1 Then
			ValidateSolicitacao = True
		Else
			ValidateSolicitacao = False
		End If
	End Function

	function validaStatusSolicitacao(ID)
		dim sql, arrsol, intsol, i
		dim statusColeta
		dim qtdCartuchos
		dim	dataReceb
		dim dataEnvioTransp
		dim dataProg
		dim dataAprov

		sql = "select status_coleta_idstatus_coleta, " & _
			  "data_recebimento, " & _
			  "data_envio_transportadora, " & _
			  "qtd_cartuchos_recebidos, " & _
			  "data_programada, " & _
			  "data_aprovacao " & _
			  "from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"'"

'		response.write sql & "<br />"

		call search(sql, arrsol, intsol)
		if intsol > -1 then
			for i=0 to intsol
				dataReceb = arrsol(1,i)
				dataEnvioTransp = arrsol(2,i)
				qtdCartuchos = arrsol(3,i)
				dataProg = arrsol(4,i)
				dataAprov = arrsol(4,i)
				statusColeta = arrsol(0,i)

				if not cint(statusColeta) = 6 then
					if isEmpty(qtdCartuchos) or isNull(qtdCartuchos) then
						if isEmpty(dataReceb) or isNull(dataReceb) then
							if (isEmpty(dataProg) or isNull(dataProg)) and (isEmpty(dataAprov) or isNull(dataAprov)) and (isEmpty(dataEnvioTransp) or isNull(dataEnvioTransp))  then
								log_error = log_error & "<b style=""color:#FF0000;"">Erro: Data Programada, Data de Aprova��o e Data Envio para Transportadora n�o est�o preenchidas / N�mero da Solicita��o ["&ID&"]</b><br />"
								validaStatusSolicitacao = false
							else
								validaStatusSolicitacao = true
							end if
						else
							log_error = log_error & "<b style=""color:#FF0000;"">Erro: Data Recebimento da Solicita��o j� foi preenchida / N�mero da Solicita��o ["&ID&"]</b><br />"
							validaStatusSolicitacao = false
						end if
					else
						log_error = log_error & "<b style=""color:#FF0000;"">Erro: Quantidade de cartuchos da Solicita��o j� foi preenchida / N�mero da Solicita��o ["&ID&"]</b><br />"
						validaStatusSolicitacao = false
					end if
				else
					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Solicita��o consta como encerrada / N�mero da Solicita��o ["&ID&"]</b><br />"
					validaStatusSolicitacao = false
				end if
			next
		else
			validaStatusSolicitacao = false
		end if
	end function

	function formatdate(data)
		dim dia
		dim mes
		dim ano
		dim arr

		if data <> "" then
			'arr = split(data, "/")
			
			'dia = arr(1)
			'mes = arr(0)
			'ano = arr(2)
					
			dia = day(data)
			mes = month(data)
			ano = year(data)

			if len(dia) = 1 then
			 dia = "0" & dia
			end if
			if len(mes) = 1 then
				mes = "0" & mes
			end if
			'formatdate = dia & "/" & mes & "/" & ano
			formatdate = ano & "-" & mes & "-" & dia '+ " 00:00:00.000"
		else
			formatdate = ""
		end if

	end function

	sub addKardex(numero_solicitacao, data_recebimento, cod_produto, qtd)
		call atualizaSolicitacao(numero_solicitacao, data_recebimento, cod_produto, qtd)
		dim sql

		if isMaster(numero_solicitacao) then
			if getIDPontoByNumSol(numero_solicitacao) > -1 then
				sql = "INSERT INTO [marketingoki2].[dbo].[Kardex]( " & _
						"[codigo_cliente], " & _
						"[data_recebimento], " & _
						"[codigo_produto], " & _
						"[descricao_produto], " & _
						"[qtd], " & _
						"[data_geracao_bonus], " & _
						"[numero_solicitacao_coleta]) " & _
						"VALUES( " & _
						""&getIDPontoByNumSol(numero_solicitacao)&", " & _
						"CONVERT(DATETIME, '"&formataDataPonto(data_recebimento)&"'), " & _
						"'"&cod_produto&"', " & _
						"'"&getDescProduto(cod_produto)&"', " & _
						""&qtd&", " & _
						"NULL, " & _
						"'"&numero_solicitacao&"')"
			end if
		else
			sql = "INSERT INTO [marketingoki2].[dbo].[Kardex]( " & _
					"[codigo_cliente], " & _
					"[data_recebimento], " & _
					"[codigo_produto], " & _
					"[descricao_produto], " & _
					"[qtd], " & _
					"[data_geracao_bonus], " & _
					"[numero_solicitacao_coleta]) " & _
					"VALUES( " & _
					""&getIDCliByNumSol(numero_solicitacao)&", " & _
					"CONVERT(DATETIME, '"&formatdate(data_recebimento)&"'), " & _
					"'"&cod_produto&"', " & _
					"'"&getDescProduto(cod_produto)&"', " & _
					""&qtd&", " & _
					"NULL, " & _
					"'"&numero_solicitacao&"')"
		end if
'			response.write sql
'			response.end
		call exec(sql)
	end sub

	function getDescProduto(id)
		dim sql, arr, intarr, i
		sql = "select descricao from produtos where idoki = '"&id&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getDescProduto = arr(0,i)
			next
		else
			getDescProduto = ""
		end if
	end function

	function getIDPontoByNumSol(num)
		dim sql, arr, intarr, i
		sql = "select a.pontos_coleta_idpontos_coleta from solicitacao_coleta_has_pontos_coleta as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.solicitacao_coleta_idsolicitacao_coleta = b.idsolicitacao_coleta " & _
				"where b.numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			getIDPontoByNumSol = arr(0,0)
		else
			getIDPontoByNumSol = -1
		end if
	end function

	function isMaster(id)
		dim sql, arr, intarr, i
		sql = "select ismaster from solicitacao_coleta where numero_solicitacao_coleta = '"&id&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cint(arr(0,i)) = 0 then
					isMaster = false
				else
					isMaster = true
				end if
			next
		end if
	end function

	function getIDCliByNumSol(num)
		dim sql, arr, intarr, i

		sql = "select c.idclientes from solicitacao_coleta as a " & _
				"left join solicitacao_coleta_has_clientes as b " & _
				"on a.idsolicitacao_coleta = b.solicitacao_coleta_idsolicitacao_coleta " & _
				"left join clientes as c " & _
				"on b.clientes_idclientes = c.idclientes " & _
				"where a.numero_solicitacao_coleta = '"&num&"'"

		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getIDCliByNumSol = arr(0,i)
			next
		else
			getIDCliByNumSol = 0
		end if
	end function

	function getIDSolicitacao(num)
		dim sql, arr, intarr, i
		sql = "select idsolicitacao_coleta from solicitacao_coleta where numero_solicitacao_coleta = '"&num&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			getIDSolicitacao = arr(0,0)
		else
			getIDSolicitacao = -1
		end if
	end function

	function validaDataReceb(data, ID)
		dim sql, arr, intarr, i
		dim valida
		sql = "select data_envio_transportadora " & _
				"from solicitacao_coleta " & _
				"where numero_solicitacao_coleta = '"&ID&"'"
		call search(sql,arr,intarr)
		if intarr > -1 then
			'Response.Write "<hr>#" & data & "#<hr>"
			'Response.Write "<hr>" & month(arr(0,i)) & "/" & day(arr(0,i)) & "/" & year(arr(0,i))&" - "&mid(data,4,2) & "/" & left(data,2) & "/" & right(data,4) & "<hr>"
			'Response.Write datediff("d",month(arr(0,i)) & "/" & day(arr(0,i)) & "/" & year(arr(0,i)), mid(data,3,2) & "/" & left(data,2) & "/" & right(data,4)) & "<hr>"
			
			'Response.Write datediff("d",month(arr(0,i)) & "/" & day(arr(0,i)) & "/" & year(arr(0,i)), mid(data,4,2) & "/" & left(data,2) & "/" & right(data,4))
			
			
			'Response.End
			
			'valida = datediff("d",month(arr(0,i)) & "/" & day(arr(0,i)) & "/" & year(arr(0,i)), mid(data,4,2) & "/" & left(data,2) & "/" & right(data,4))
			valida = datediff("d", day(arr(0,i)) & "/" &  month(arr(0,i)) & "/" & year(arr(0,i)), day(data) & "/" & month(data) & "/" & year(data))
			
			
			if valida >= 0 then
				validaDataReceb = true
			else
				validaDataReceb = false
			end if
		else
			validaDataReceb = false
		end if
	end function

%>
<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
<script>
	function validate() {
		if (document.frmKardex.attach1.value == "") {
			alert("Escolha um arquivo para exportar!");
			return;
		} else {
			document.frmKardex.submit();
		}
	}
</script>
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<form action="" name="frmKardex" method="POST" enctype="multipart/form-data">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableCadCliente">
						<tr>
							<td id="explaintitle" align="center" colspan="2">Importa��o de arquivo texto para alimenta��o do Kardex</td>
						</tr>
						<tr>
							<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td width="40%" align="right"><b>Arquivo texto a ser importado:</b> </td>
							<td align="left"><input type="file" class="btnform" name="attach1" size="35" />&nbsp;</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td align="center"><input type="button" class="btnform" name="enviar" value="Importar" onClick="validate()" /></td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="2" align="center">
							<%Call Submit()%>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="2">
								<table cellspacing="1" cellpadding="1" width="100%">
									<tr>
										<td id="explaintitle" align="center">Lista de Atualiza��es</td>
									</tr>
								</table>
								<tr>
									<td colspan="2">
										<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro">
											<tr>
												<th>N� Sol de Coleta</th>
												<th>Data Recebimento</th>
												<th>Cod do Produto</th>
												<th>Qtd do Produto</th>
											</tr>
											<%=htmlFileImprimir%>
										</table>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro">
											<tr>
												<th>Log de Erros</th>
											</tr>
											<tr>
												<td><%=log_error%></td>
											</tr>
										</table>
									</td>
								</tr>
							</td>
						</tr>
					</table>
					</form>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
