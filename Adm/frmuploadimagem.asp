<!--#include file="../_config/_config.asp" -->
<!-- #include file="../_config/freeaspupload.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Response.Expires = -1
	Server.ScriptTimeout = 600

	dim uploadsDirVar
	dim FileUploaded

	If Left(Request.ServerVariables("LOCAL_ADDR"),3) = "127" Or Left(Request.ServerVariables("LOCAL_ADDR"),3) = "192" Then
	  uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Home\"
	Else
	  uploadsDirVar = Request.ServerVariables("APPL_PHYSICAL_PATH") & "adm\Home\"
	End If

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

	function getListFiles()
		dim files
		dim fso
		dim folder
		dim ofolder
		dim nome_file
		dim i
		dim style
		dim nome
		dim datecreatedfile
		dim txt

		i = 0
		folder = request.servervariables("APPL_PHYSICAL_PATH")&"adm/home/"
		set fso = server.createobject("scripting.filesystemobject")
		set ofolder = fso.getfolder(folder)
		set files = ofolder.files
		for each nome_file in files
			if i mod 2 = 0 then
				style = "class=""classColorRelPar"""
			else
				style = "class=""classColorRelImpar"""
			end if
			i = i + 1
			datecreatedfile = nome_file.dateCreated
			nome = split(nome_file, "\")
			'peterson aquino 25-5-2014
			txt = nome_file.Name

				getListFiles = getListFiles & "<tr>"
				if left(request.servervariables("LOCAL_ADDR"),3) = "127" then
					'getListFiles = getListFiles & "<td "&style&"><b>"&nome(5)&"</b> / "&nome(6)&" <a href=""http://localhost:81/sgrs/adm/home/"&nome(6)&""">clique aqui</a></td><td "&style&">"&formatdatetime(datecreatedfile, 2)&"</td>"
					getListFiles = getListFiles & "<td "&style&"><b>" & txt & "<a href=""http://localhost:81/sgrs/adm/home/" & txt & """>clique aqui</a></td><td "&style&">"&formatdatetime(datecreatedfile, 2)&"</td>"
				else
					'getListFiles = getListFiles & "<td "&style&"><b>"&nome(5)&"</b> / "&nome(6)&" <a href=""http://www.sustentabilidadeoki.com.br/adm/home/"&nome(6)&""">clique aqui</a></td><td "&style&">"&DateRight(formatdatetime(datecreatedfile, 2))&"</td>"
					getListFiles = getListFiles & "<td "&style&"><b>" & txt & "<a href=""http://www.sustentabilidadeoki.com.br/adm/home/"&txt&""">clique aqui</a></td><td "&style&">"&DateRight(formatdatetime(datecreatedfile, 2))&"</td>"
				end if
				getListFiles = getListFiles & "</tr>"
		next

		set fso = nothing
		set ofolder = nothing
		set folder = nothing
	end function

	sub deleteFiles()
		dim files
		dim fso
		dim file_name
		dim cont_files
		dim folder
		set fso = server.createobject("scripting.filesystemobject")
		files = split(replace(request.form("bannerfile"), " ", ""), ",")
		response.write request.form("bannerfile") & "<br />"
		for cont_files=0 to ubound(files)
			folder = server.MapPath("home/") & "\"
'			response.write files(cont_files) & "<br />"
			folder = folder & files(cont_files)
'			response.write folder
			set file_name = fso.getfile(folder)
			file_name.delete
		next

		set file_name = nothing
		set fso = nothing
	end sub

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
						<table cellpadding="1" cellspacing="1" width="100%" id="tableCadCliente">
							<form action="" name="frmEletronicFileTransp" method="POST" enctype="multipart/form-data">
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
								<td width="40%" align="right"><b>Arquivo para o Banner a ser importado:</b> </td>
								<td align="left"><input type="file" style="" name="attach1" size="35" />&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2">&nbsp;</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align="center">
									<input type="button" style="" name="enviar" value="Importar" onClick="validate()" />&nbsp;
								</td>
							</tr>
							<tr>
								<td colspan="2" align="center">
								<%Call Submit()%>
								</td>
							</tr>
							<tr>
								<td colspan="12">
									<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro">
										<tr>
											<th width="70%">Nome</th>
											<th>Data Criação</th>
										</tr>
										<%=getListFiles()%>
									</table>
								</td>
							</tr>
							</form>
					</table>
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
