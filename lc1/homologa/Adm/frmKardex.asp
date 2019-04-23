<!--#include file="../_config/_config.asp" -->
<!-- #include file="../_config/freeaspupload.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Response.Expires = -1
	Server.ScriptTimeout = 600
	Session.Timeout = 600

	Dim uploadsDirVar
	Dim diagnostics
	Dim FileUploaded
	Dim QtdRealProduto
	dim log_error
	dim berror
	
	log_error = ""
	berror = false
	
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
					SaveFiles = "<B>Arquivo Exportado:</B> "
					For Each fileKey in Upload.UploadedFiles.keys
							SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
							FileUploaded = Upload.UploadedFiles(fileKey).FileName
					next
			Else
					SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
			End If
	End Function
	
	Sub Submit()
		If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
				diagnostics = TestEnvironment()
				If diagnostics <> "" Then
						Response.Write diagnostics
				End If
		Else
				Response.Write SaveFiles()
		End If
	End Sub
	
	Function LerArquivo()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			On Error Resume Next
			Dim oFso
			Dim oFile
			Dim oStream
			Dim Linha
			Dim Ret
			Dim Arr
			Dim i
			Dim Cont
			Dim NumSolicitacao
			Dim contTimeWhile
			
			contTimeWhile = 0
			NumSolicitacao = ""
			Set oFso = Server.CreateObject("Scripting.FileSystemObject")
			Set oFile = oFso.GetFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Kardex\Domiciliar\" & FileUploaded)
			Set oStream = oFile.OpenAsTextStream(1,False) 
			
			While Not oStream.AtEndOfStream
				contTimeWhile = contTimeWhile + 1
				Linha = oStream.ReadLine
				Arr = Split(Linha,";")
'				response.write berror & "<br />"
				If ValidateSolicitacao(Arr(0)) Then
					If ValidateProduct(Arr(2)) Then
'						response.write validaStatusSolicitacao(Arr(0), contTimeWhile) & "<br />"
						if validaStatusSolicitacao(Arr(0), contTimeWhile) then
							If contTimeWhile = 1 Then
								NumSolicitacao = Arr(0)
							Else
								If NumSolicitacao <> Arr(0) Then
									NumSolicitacao = Arr(0)
									QtdRealProduto = 0
									berror = false
								End If
							End If
							Ret = Ret & "<tr>"
							For i=0 To Ubound(Arr)
								If Cont Mod 2 = 0 Then
									Ret = Ret &  "<td class='classColorRelPar'>" & Arr(i) & "</td>"
								Else
									Ret = Ret &  "<td class='classColorRelImpar'>" & Arr(i) & "</td>"
								End If	
							Next
							Ret = Ret & "</tr>"
							Cont = Cont + 1
							Call AtualizaSol(Arr(0), Arr(1), Arr(2), Arr(3), contTimeWhile)
						else
							berror = true
							log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de atualização de status não pode ser efetuado - Solicitação: ["&Arr(0)&"]</b><br />"	
						end if						
					Else
						berror = true
						log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Produto inexistente / Produto: "&Arr(2)&"</b><br />"
					End If	
				Else
					berror = true
					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Solicitação inexistente / Solicitação: "&Arr(0)&"</b><br />"
				End If
			Wend
			
			oStream.Close
			Set oFso = Nothing
			QtdRealProduto = 0
			LerArquivo = Ret
			If Error <> 0 Then
				Response.Write "<tr><td colspan=""5"">Erro na operação de Atualização</td></tr>"
				Exit Function
			End If 
		End If	
	End Function
	
	Sub AtualizaSol(NumSol, DataReceb, CodProduto, QtdProduto, cont)
		Dim sSql, arrSol, intSol, i
		Dim arrSol2, intSol2, j
		dim sql, arr, intarr
		QtdRealProduto = CInt(QtdRealProduto) + CInt(QtdProduto)
		
		if DataReceb <> "" and not berror then
			sql = "select * from solicitacao_coleta where data_recebimento = convert(datetime,'"&DataReceb&"') and [numero_solicitacao_coleta] = '" &NumSol& "'"
'			response.write sql
'			response.end
			call search(sql, arr, intarr)
			if intarr > -1  and cont = 1then
				berror = true
				log_error = log_error & "<b style=""color:#FF0000;"">A Solicitação "&NumSol&" já foi operacionada!</b><br />"
			else	
				if validaDataReceb(DataReceb, NumSol) then
					sSql = "SELECT [idSolicitacao_coleta] " & _
									"FROM [marketingoki2].[dbo].[Solicitacao_coleta] " & _
									"WHERE [numero_solicitacao_coleta] = '" &NumSol& "' " & _
									"AND MONTH(data_solicitacao) = " & Month(Now())
'					response.write sSql
'					response.end										
					Call search(sSql, arrSol, intSol)
					If intSol > -1 Then
						For i=0 To intSol
							sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
											"SET " & _ 
											"[Status_coleta_idStatus_coleta] = 6, " & _ 
											"[qtd_cartuchos_recebidos] = "&QtdRealProduto&", " & _ 
											"[data_recebimento] = '"&DataReceb&"' " & _ 
											"WHERE [idSolicitacao_coleta] = "&arrSol(0,i)
							Call exec(sSql)	
							sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
											"[Produtos_idProdutos], " & _ 
											"[Solicitacao_coleta_idSolicitacoes_coleta], " & _ 
											"[quantidade]) " & _
											"VALUES( " & _
											"'"&CodProduto&"', " & _ 
											""&arrSol(0,i)&", " & _ 
											""&QtdProduto&")"
							Call exec(sSql)				
							call addKardex(DataReceb, NumSol, CodProduto, QtdProduto)				
						Next
					End If
				else
					berror = true
					log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Data incorreto / Solicitação ["&NumSol&"]</b><br />"	
				end if	
			end if	
		else
			berror = true
			log_error = log_error & "<b style=""color:#FF0000;"">Erro: Processamento de Data incorreto / Solicitação ["&NumSol&"]</b><br />"	
		end if	
	End Sub
	
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
		sSql = "select * from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"'"
'		response.write sSql
'		response.end
		Call search(sSql, arrSol, intSol)
		If intSol > -1 Then
			ValidateSolicitacao = True
		Else
			ValidateSolicitacao = False		
		End If
	End Function
	
	function validaStatusSolicitacao(ID, vezes)
'		response.write vezes & "<br />"
		if vezes = 1 and not berror then
			dim sql, arrsol, intsol, i
			sql = "select status_coleta_idstatus_coleta from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"' and status_coleta_idstatus_coleta <> 6"
			call search(sql, arrsol, intsol)
			if intsol > -1 then
				for i=0 to intsol
					if cint(arrsol(0,i)) = 7 or cint(arrsol(0,i)) = 8 then
						validaStatusSolicitacao = true	
					else
						validaStatusSolicitacao = false
					end if
				next
			else	
				validaStatusSolicitacao = false
			end if
		else
			if not berror then
				validaStatusSolicitacao = true
			else
				validaStatusSolicitacao = false
			end if	
		end if	
	end function 
	
	function formatdate(data)
		dim dia
		dim mes
		dim ano
		dim arr

		arr = split(data, "/")

		dia = arr(1)
		mes = arr(0)
		ano = arr(2)
		
		if len(dia) = 1 then
		 dia = "0" & dia
		end if
		if len(mes) = 1 then
			mes = "0" & mes
		end if 
		
		formatdate = dia & "/" & mes & "/" & ano
	end function
	
	sub addKardex(data_recebimento, numero_solicitacao, cod_produto, qtd)
		dim sql
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
				"CONVERT(DATETIME, '"&data_recebimento&"'), " & _ 
				"'"&cod_produto&"', " & _ 
				"'"&getDescProduto(cod_produto)&"', " & _ 
				""&qtd&", " & _ 
				"NULL, " & _ 
				"'"&numero_solicitacao&"')"
		'response.Write sql
		'response.End					
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
	
	function validaDataReceb(data, ID)
		dim sql, arr, intarr, i
		dim valida
		sql = "select data_programada " & _
				"from solicitacao_coleta " & _ 
				"where numero_solicitacao_coleta = '"&ID&"'"
		call search(sql,arr,intarr)
		if intarr > -1 then	
			valida = datediff("d",arr(0,i),formatdatetime(data,2))
			response.write valida & "<br />"
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
					<form action="frmKardex.asp" name="frmKardex" method="POST" enctype="multipart/form-data">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableCadCliente">
						<tr>
							<td id="explaintitle" align="center" colspan="2">Importação de arquivo texto para alimentação do Kardex</td>
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
										<td id="explaintitle" align="center">Lista de Atualizações</td>
									</tr>
								</table>
								<tr>
									<td colspan="2">
										<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro">
											<tr>
												<th>N° Sol de Coleta</th>
												<th>Data Recebimento</th>
												<th>Cod do Produto</th>
												<th>Qtd do Produto</th>
											</tr>
											<%=LerArquivo%>
										</table>	
									</td>
								</tr>
								<tr>
									<td colspan="2"><%=log_error%></td>
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
