<!--#include file="../_config/_config.asp" -->
<%
'|--------------------------------------------------------------------
'| Arquivo: frmCadCliente.asp																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação : 15/04/2007																		 
'| Descrição: Arquivo de Formulário para cadastro de Cliente
'|--------------------------------------------------------------------
%>
<%Call open()%>
<% Response.Charset="ISO-8859-1" %>
<%	
	Dim method
	Dim id
	
	method = Request.QueryString("sub")
	id = Request.QueryString("id")
	
	If method = "getminimo" Then
		Call getMinimo(id)
	Elseif method = "getEndColeta" Then
		Call getEndColeta(id)
	ElseIf method = "gettypecoleta" Then
		Call getTypeColeta(id)
	ElseIf method = "getcependereco" Then
		Call getCepEndereco(id)
	ElseIf method = "getlistpontocoleta" Then
		Call getListPontoColeta(id)
	ElseIf method = "getcheckusercontato" Then
		Call getCheckUserContato(Request.QueryString("user"), Request.QueryString("senha"))
	ElseIf method = "getcheckcnpjempresa" Then
		Call getCheckCNPJEmpresa(id)
	ElseIf method = "getcheckcpf" Then
		Call getCheckCPF(id)
	elseif method = "getcheckusuario"	then
		call getCheckUsuario(Request.QueryString("user"))
	End If
	
	Sub getMinimo(categoria)
		Dim sSql, arrMinimo, intMinimo, i
		sSql = "SELECT [minCartuchos] " & _
			   "FROM [marketingoki2].[dbo].[Categorias] " & _
			   "WHERE [idCategorias] = " & categoria
		Call search(sSql, arrMinimo, intMinimo)
		If intMinimo > -1 Then
			For i=0 To intMinimo
				Response.Write arrMinimo(0,i)
			Next
		End If	   
	End Sub
	
	Sub getTypeColeta(categoria)
		Dim sSql, arrTypeColeta, intTypeColeta, i
		sSql = "SELECT [isColetaDomiciliar] " & _
			   "FROM [marketingoki2].[dbo].[Categorias] " & _
			   "WHERE [idCategorias] = " & categoria
		Call search(sSql, arrTypeColeta, intTypeColeta)
		If intTypeColeta > -1 Then
			For i=0 To intTypeColeta
				Response.Write arrTypeColeta(0,i)
			Next
		End If	   
	End Sub
	
	Sub getEndColeta(id)
		Dim sSql, arrCep2, intCep, i, xmlWrite
		sSql = 	"select top 1 '0' as id, A.cep, A.logradouro, A.bairro, A.municipio, A.estado, B.numero_endereco_coleta, B.compl_endereco_coleta " & _
				"FROM cep_consulta_has_Clientes as A " & _
				"inner join Clientes as B on A.Clientes_idClientes = B.idClientes " & _
				"WHERE Clientes_idClientes = " & id & " and A.isEnderecoComum = 0 " 

			Call search(sSql, arrCep2, intCep)
			With Response
				If intCep > -1 Then
					For i=0 To intCep
						.Write arrCep2(0,i)&";"&arrCep2(1,i)&";"&arrCep2(2,i)&";"&arrCep2(3,i)&";"&arrCep2(4,i)&";"&arrCep2(5,i)&";"&arrCep2(6,i)&";"&arrCep2(7,i)
					Next
				End If
			End With
	End Sub
	
	Sub getCepEndereco(cep)
		Dim sSql, arrCep, intCep, i, xmlWrite
		sSql = "SELECT top 1 [idcep_consulta], " & _
					 "[cep], " & _ 
					 "[logradouro], " & _ 
					 "[bairro], " & _ 
					 "[municipio], " & _ 
					 "[estado], " & _ 
					 "rtrim([tipologradouro]) " & _ 
					 "FROM [marketingoki2].[dbo].[cep_consulta] " & _
					 "WHERE [cep] = " & cep
		Call search(sSql, arrCep, intCep)
		With Response
			If intCep > -1 Then
				For i=0 To intCep
					.Write arrCep(0,i)&";"&arrCep(1,i)&";"&arrCep(2,i)&";"&arrCep(3,i)&";"&arrCep(4,i)&";"&arrCep(5,i)&";"&arrCep(6,i)					
				Next
			End If
		End With
	End Sub

	Sub getListPontoColeta(cep)
		Dim sSql, arrPontoColeta, intPontoColeta, i
		Dim inicio, fim
		Dim tableRet
		Dim arrEstado, intEstado, iEs, estado
		Dim arrPontoColeta2, intPontoColeta2, iP
		Dim sqlCep, sqlPontoEstado

		inicio = Left(cep, 3) & "00000"
		fim = Left(cep, 3) & "99999"

'		sSql = "SELECT * FROM Pontos_coleta AS A " & _
'				"WHERE A.status_pontocoleta = 1 AND A.cep BETWEEN '"&inicio&"' AND '"&fim&"'"
					 
'		sSql = "SELECT " & _
'						"a.razao_social, " & _ 
'						"b.tipologradouro, " & _
'						"b.logradouro, " & _ 
'						"a.complemento_endereco, " & _
'						"a.numero_endereco, " & _
'						"b.bairro, " & _
'						"b.municipio, " & _
'						"b.estado, " & _
'						"b.cep, " & _  
'						"A.idpontos_coleta " & _
'						"FROM Pontos_coleta AS A " & _ 
'						"LEFT JOIN cep_consulta AS B " & _ 
'						"ON A.cep_consulta_idcep_consulta = B.idcep_consulta " & _ 
'						"WHERE A.status_pontocoleta = 1 AND B.cep BETWEEN '"&inicio&"' AND '"&fim&"'"
						
		sSql = "select " & _
				"idPontos_coleta, " & _
				"razao_social, " & _
				"nome_fantasia, " & _
				"cnpj, " & _
				"logradouro, " & _
				"numero_endereco, " & _
				"complemento_endereco, " & _
				"bairro, " & _
				"cep, " & _
				"municipio, " & _
				"estado " & _
				"from pontos_coleta " & _
				"where status_pontocoleta = 1 and cep between '"&inicio&"' AND '"&fim&"'"
		
		' -- consultaas dos campos
		'idponto coleta				= 0
		'razao social				= 1 		

		'nome fantasia				= 2
		'cnpj						= 3
		'logradouro					= 4
		'numero endereco			= 5
		'complemento endereco		= 6
		'bairro						= 7
		'cep						= 8
		'municipio					= 9
		'estado						= 10
					 
		Call search(sSql, arrPontoColeta, intPontoColeta)
		tableRet = "<table cellpadding='3' cellspacing='4' width='100%' id='tableListPontoColeta' align='center'>"
				tableRet = tableRet & "<tr>"
				tableRet = tableRet & "<th><img src='img/check.gif' alt='Selecionar' /></th>"
				tableRet = tableRet & "<th>Razão Social</th>"
				tableRet = tableRet & "<th>Endereço</th>"
				tableRet = tableRet & "<th>Bairro</th>"
				tableRet = tableRet & "<th>Cidade</th>"
				tableRet = tableRet & "<th>Estado</th>"
				tableRet = tableRet & "<th>Cep</th>"
				tableRet = tableRet & "</tr>"
		If intPontoColeta > -1 Then
			For i=0 To intPontoColeta
				tableRet = tableRet & "<tr>"
				tableRet = tableRet & "<td><input type='radio' name='radioCheckPonto' id='radioCheckPonto"&i&"' value='"&arrPontoColeta(0,i)&"' /></td>"
				tableRet = tableRet & "<td>"&arrPontoColeta(1,i)&"</td>"
				tableRet = tableRet & "<td>"&trim(arrPontoColeta(4,i))&" n° "&trim(arrPontoColeta(5,i))&"</td>"
				tableRet = tableRet & "<td>"&trim(arrPontoColeta(7,i))&"</td>"
				tableRet = tableRet & "<td>"&trim(arrPontoColeta(9,i))&"</td>"
				tableRet = tableRet & "<td>"&trim(arrPontoColeta(10,i))&"</td>"
				tableRet = tableRet & "<td>"&trim(arrPontoColeta(8,i))&"</td>"
				tableRet = tableRet & "</tr>"
			Next
		Else
			sqlCep = "SELECT estado FROM cep_consulta WHERE cep = '"&cep&"'"
			Call search(sqlCep, arrEstado, intEstado)
			If intEstado > -1 Then
				For iEs=0 To intEstado
					estado = arrEstado(0,iEs)
				Next
				
				sqlPontoEstado = "select " & _
						"idPontos_coleta, " & _
						"razao_social, " & _
						"nome_fantasia, " & _
						"cnpj, " & _
						"logradouro, " & _
						"numero_endereco, " & _
						"complemento_endereco, " & _
						"bairro, " & _
						"cep, " & _
						"municipio, " & _
						"estado " & _
						"from pontos_coleta " & _
						"where status_pontocoleta = 1 and estado = '" & estado & "'"				
						
				' -- consultaas dos campos
				'idponto coleta				= 0
				'razao social				= 1 		
				'nome fantasia				= 2
				'cnpj						= 3
				'logradouro					= 4
				'numero endereco			= 5
				'complemento endereco		= 6
				'bairro						= 7
				'cep						= 8
				'municipio					= 9
				'estado						= 10
				
				Call search(sqlPontoEstado, arrPontoColeta2, intPontoColeta2)				
				If intPontoColeta2 > -1 Then
					For iP=0 To intPontoColeta2
						tableRet = tableRet & "<tr>"
						tableRet = tableRet & "<td><input type='radio' name='radioCheckPonto' id='radioCheckPonto"&iP&"' value='"&arrPontoColeta2(0,iP)&"' /></td>"
						tableRet = tableRet & "<td>"&arrPontoColeta2(1,iP)&"</td>"
						tableRet = tableRet & "<td>"&trim(arrPontoColeta2(4,iP))&" n° "&trim(arrPontoColeta2(5,iP))&"</td>"
						tableRet = tableRet & "<td>"&trim(arrPontoColeta2(7,iP))&"</td>"
						tableRet = tableRet & "<td>"&trim(arrPontoColeta2(9,iP))&"</td>"
						tableRet = tableRet & "<td>"&trim(arrPontoColeta2(10,iP))&"</td>"
						tableRet = tableRet & "<td>"&trim(arrPontoColeta2(8,iP))&"</td>"
						tableRet = tableRet & "</tr>"
					Next
					intPontoColeta = intPontoColeta2
				Else
					tableRet = tableRet & "<tr><td colspan='7'>Nenhum Ponto de coleta mais próximo ao seu endereço.</td></tr>"
				End If
			Else	
				tableRet = tableRet & "<tr><td colspan='7'>Nenhum Ponto de coleta mais próximo ao seu endereço.</td></tr>"
			End If
		End If
		tableRet = tableRet & "</table>"
		Response.Write tableRet&";"& intPontoColeta
	End Sub

	Sub getCheckUserContato(User, Senha)
		Dim sSql, arrUser, intUser
		dim sql, arr, intarr

		sSql = "SELECT " & _ 
						"[idContatos] " & _ 
						"FROM [marketingoki2].[dbo].[Contatos] " & _
						"WHERE [usuario] = '"&User&"' " & _ 
						"AND [senha] = '"&Senha&"'"

		Call search(sSql, arrUser, intUser)

		If intUser > -1 Then
			Response.Write "true"
		Else
			sql = "SELECT " & _ 
						"[idContatos] " & _ 
						"FROM [marketingoki2].[dbo].[Contatos] " & _
						"WHERE [senha] = '"&Senha&"'"
			call search(sql, arr, intarr)
			if intarr > -1 then			
				Response.Write "true"
			else	
				Response.Write "false"
			end if	
		End If
	End Sub
	
	sub getCheckUsuario(user)
		Dim sql2, arrUser2, intUser2
			sql2 = "SELECT " & _ 
							"[idContatos] " & _ 
							"FROM [marketingoki2].[dbo].[Contatos] " & _
							"WHERE [usuario] = '"&user&"' " 
			call search(sql2, arrUser2, intUser2)
			if intUser2 > -1 then
				response.write "true"
			else
				response.write "false"
			end if	
	end sub

	Sub getCheckCNPJEmpresa(CNPJ)
		Dim sSql, arrCNPJ, intCNPJ
		Dim sSql2, arrCNPJ2, intCNPJ2
		Dim Ret
		
		If Session("IDCliente") <> "" Then
			sSql2 = "SELECT " & _ 
							"[idClientes] " & _
							"FROM [marketingoki2].[dbo].[Clientes] " & _
							"WHERE [cnpj] = '"&CNPJ&"' " & _
							"AND [idClientes] = " & Session("IDCliente")
	
			Call search(sSql2, arrCNPJ2, intCNPJ2)
		Else	
			intCNPJ2 = -1			
		End If
		sSql = "SELECT " & _ 
						"[idClientes] " & _
						"FROM [marketingoki2].[dbo].[Clientes] " & _
						"WHERE [cnpj] = '"&CNPJ&"'"
		Call search(sSql, arrCNPJ, intCNPJ)

		If intCNPJ > -1 And  intCNPJ2 = -1 Then
			Response.Write "true"
		Else
			Response.Write "false"
		End If
	End Sub
	
	Sub getCheckCPF(CPF)
		Dim sSql, arrCPF, intCPF
		Dim sSql2, arrCPF2, intCPF2
		Dim Ret
		
		If Session("IDCliente") <> "" Then
			sSql2 = "SELECT " & _ 
							"[idClientes] " & _
							"FROM [marketingoki2].[dbo].[Clientes] " & _
							"WHERE [cnpj] = '"&CPF&"' " & _
							"AND [idClientes] = " & Session("IDCliente")
	
			Call search(sSql2, arrCPF2, intCPF2)
		Else	
			intCPF2 = -1			
		End If

		sSql = "SELECT " & _ 
						"[idClientes] " & _
						"FROM [marketingoki2].[dbo].[Clientes] " & _
						"WHERE [cnpj] = '"&CPF&"'"
		Call search(sSql, arrCPF, intCPF)

		If intCPF > -1 And  intCPF2 = -1 Then
			Response.Write "true"
		Else
			Response.Write "false"
		End If
		
	End Sub
	
%>
<%Call close()%>