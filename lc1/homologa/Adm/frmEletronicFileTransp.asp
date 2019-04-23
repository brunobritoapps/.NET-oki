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
	dim log_error

	If Left(Request.ServerVariables("LOCAL_ADDR"),3) = "127" Or Left(Request.ServerVariables("LOCAL_ADDR"),3) = "192" Then
	  'uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Transportadora\"
	  uploadsDirVar = server.MapPath("Transportadora")
	Else
	  'uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Transportadora\"
	  uploadsDirVar = server.MapPath("Transportadora")
	End If
	'Response.Write uploadsDirVar & "<hr>"
	
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

		'If something fails inside the script, but the exception is handled
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
			Dim i, x
			Dim Cont
			Dim sStyle

			Set oFso = Server.CreateObject("Scripting.FileSystemObject")
			Set oFile = oFso.GetFile(uploadsDirVar & "\" & FileUploaded)
			Set oStream = oFile.OpenAsTextStream(1,False)

			oConn.BeginTrans

			Do While Not oStream.AtEndOfStream
				Linha = oStream.ReadLine
				'Response.Write "<hr>" & Linha & "<br>"
				'Response.Write "#$" & len(trim(linha)) & "$#<br>"
				'Response.End
				
				Arr = Split(Linha,";")
				Ret = Ret & "<tr>"
				
				if len(trim(linha)) > 0 then
					for each item in Arr
						x = x + 1
					next						
					if x <> 4 then
						log_error = log_error & "<tr><td colspan=""4""><b style=""color:#FF0000;"">Erro: Solicitação não pôde ser processada - O Arquivo está incompleto.</b></td></tr>"
					end if
					x = 0
				end if
					
				If ValidateSolicitacao(Arr(2)) Then
					For i=0 To Ubound(Arr)
						If Cont Mod 2 = 0 Then
							sStyle = "classColorRelPar"
						else
							sStyle = "classColorRelImpar"
						end if

						if i = 0 then
							Ret = Ret &  "<td class='"&sStyle&"'>" & getNomeTransp(montaCnpj(Arr(i))) & "</td>"
						else
							if i = 3 then
								Ret = Ret &  "<td class='"&sStyle&"'>" & montaData(Arr(i)) & "</td>"
							else
								Ret = Ret &  "<td class='"&sStyle&"'>" & Arr(i) & "</td>"
							end if
						end if
					Next
					Ret = Ret & "</tr>"
					
					'Response.Write montaCnpj(Arr(0)) &" # "& Arr(1) &" # "&  Arr(2) &" # "&  montaData(Arr(3)) &"<hr>"
					
					Call AtualizaSol(montaCnpj(Arr(0)), Arr(1), Arr(2), montaData(Arr(3)))
					Cont = Cont + 1
				Else
					log_error = log_error & "<tr><td colspan=""4""><b style=""color:#FF0000;"">Erro: Solicitação não pôde ser processada - Solicitação:["&Arr(2)&"] Inexistente ou não aprovada</b></td></tr>"
				End If				
			Loop

			oStream.Close
			Set oFso = Nothing
			if log_error <> "" then
				oConn.RollbackTrans
				LerArquivo = log_error & "<tr><td colspan=""4""><b style=""color:#FF0000;"">Atenção, nenhuma atualização foi efetuada. Altere o arquivo e tente novamente.</b></td></tr>"
			else
				oConn.CommitTrans
				LerArquivo = Ret
			end if
			If Error <> 0 Then
				Response.Write "<tr><td colspan=""5"">Erro na operação de Atualização</td></tr>"
				oConn.RollbackTrans
				Exit Function
			else
				oConn.CommitTrans				
			End If
		End If
	End Function

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

	function getIDTranspByCNPJ(cnpj)
		dim sql, arr, intarr, i
		sql = "select idtransportadoras from transportadoras where cnpj = '"&cnpj&"'"
'		response.write sql
'		response.end
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getIDTranspByCNPJ = arr(0,i)
			next
		else
			getIDTranspByCNPJ = -1
		end if
	end function

	function montaData(data)
		dim dia
		dim mes
		dim ano

		dia = left(data,2)
			
			if len(dia) = 1 then
				dia = "0"&dia
			end if
			
		mes = mid(data,3,2)
			
			if len(mes) = 1 then
				mes = "0"&mes
			end if		
		
		ano = right(data,4)
			if len(ano) = 2 then
				ano = "20"&ano
			end if
		
		
		montaData = dia&"/"&mes&"/"&ano
	end function

	function montaCnpj(cnpj)
		dim retorno
		dim digs1
		dim digs2
		dim digs3
		dim digs4
		dim digs5
		retorno = cnpj
		digs1 = left(retorno,2)
		digs2 = mid(retorno,3,3)
		digs3 = mid(retorno,6,3)
		digs4 = mid(retorno,9,4)
		digs5 = right(retorno,2)
		montaCnpj = digs1&"."&digs2&"."&digs3&"/"&digs4&"-"&digs5
	end function


	Sub AtualizaSol(CodTransp, NumRec, NumSol, DataProg)
		'dataProg vem como dd/mm/yyyy
		Dim sSql, arrSol, intSol, i
		Dim sSql2
		sSql = "select a.solicitacao_coleta_idsolicitacao_coleta, " & _
						"b.data_envio_transportadora " & _
						"from solicitacao_coleta_has_transportadoras as a " & _
						"left join solicitacao_coleta as b " & _
						"on a.Solicitacao_coleta_idSolicitacao_coleta = b.idSolicitacao_coleta " & _
						"where b.numero_solicitacao_coleta = '"&NumSol&"' and a.transportadoras_idtransportadoras = "&getIDTranspByCNPJ(CodTransp)
		'response.write sSql & "<br>"
		'response.end
		Call search(sSql, arrSol, intSol)
		If intSol > -1 Then
			For i=0 To intSol
				if not validaData(day(arrSol(1,i)) & "/" & month(arrSol(1,i)) & "/" & year(arrSol(1,i)), DataProg) then
					sSql2 = "update solicitacao_coleta_has_transportadoras " & _
									"set numero_reconhecimento_transportadora = '"&NumRec&"' " & _
									"where transportadoras_idtransportadoras = "&getIDTranspByCNPJ(CodTransp)&" and solicitacao_coleta_idsolicitacao_coleta = '"&arrSol(0,i)&"'"
					'response.write sSql2 & "<br />"
					Call exec(sSql2)
					
					sSql2 = "update solicitacao_coleta set status_coleta_idstatus_coleta = 7, data_programada = convert(datetime, '"&right(dataProg,4) & "-" & mid(dataProg,4,2) & "-" & left(dataProg,2)&"') where idsolicitacao_coleta = " & arrSol(0,i)
					'sSql2 = "update solicitacao_coleta set status_coleta_idstatus_coleta = 7, data_programada = convert(datetime, '"&DataProg&"') where idsolicitacao_coleta = " & arrSol(0,i)
					
					'response.write sSql2 & "<hr>"
					'response.end
					Call exec(sSql2)
				else
					log_error = log_error & "<tr><td colspan=""4""><b style=""color:#FF0000;"">Erro: Data Programada menor que a Data de Envio á Transportadora</b></td></tr>"
				end if
			Next
		End If
	End Sub

	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia
		Dia = Left(sDate, 2)
		Mes = Mid(sDate, 4, 2)
		Mes = Replace(Mes, "/" ,"")
		If Len(Mes) = 1 Then
			Mes = "0" & Mes
		End If
		Ano = Right(sDate, 4)

		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function
	
	Function ValidateSolicitacao(ID)
		Dim sSql, arrSol, intSol, i
		
		sSql = "select * from solicitacao_coleta where numero_solicitacao_coleta = '"&ID&"' and status_coleta_idstatus_coleta = 5"
		'response.write sSql & "<br />"
		'Response.End
		Call search(sSql, arrSol, intSol)
		If intSol > -1 Then
			ValidateSolicitacao = True
		Else
			ValidateSolicitacao = False
		End If
	End Function

	function getNomeTransp(id)
		dim sql, arr, intarr, i
		sql = "select razao_social from transportadoras where cnpj = '"&id&"'"
'		response.write sql
'		response.end
		call search(sql ,arr, intarr)
		if intarr > -1 then
			getNomeTransp = arr(0,0)
		else
			getNomeTransp = ""
		end if
	end function

	function validaData(dataEnvio, dataProg)
		'dataProg vem como dd/mm/yyyy
		'Response.Write dataEnvio & "<br>"
		'Response.Write dataProg & "<br>"
		'Response.Write datediff("d", formatdatetime("7/4/2008",2), formatdatetime(dataProg,2))		
		'Response.End
		
		
		dim valida
		'valida = datediff("d", dataEnvio, mid(dataProg,4,2) & "/" & left(dataProg,2) & "/" & right(dataProg,4))
		valida = datediff("d", dataEnvio, dataProg)
'		response.write "valida: " & valida & "<br />"
'		response.write "dataEnvio: " & dataEnvio & "<br />"
'		response.write "dataProg: " & formatdatetime(dataProg, 2) & "<br />"
		if valida >= 0 then
			validaData = false
		else
			validaData = true
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
		if (document.frmEletronicFileTransp.attach1.value == "") {
			alert("Escolha um arquivo para exportar!");
			return;
		} else {
			document.frmEletronicFileTransp.submit();
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
					<form action="frmEletronicFileTransp.asp" name="frmEletronicFileTransp" method="POST" enctype="multipart/form-data">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableCadCliente">
						<tr>
							<td id="explaintitle" align="center" colspan="2">Importação de arquivo eletrônico da Transportadora</td>
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
												<th>Transportadora</th>
												<th>N°. conhecimento da Transportadora</th>
												<th>N°. da Solicitação de Coleta</th>
												<th>Data Programada</th>
											</tr>
											<%=LerArquivo()%>
										</table>
									</td>
								</tr>
								<tr>
									<td colspan="2"><%'= log_error %></td>
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
