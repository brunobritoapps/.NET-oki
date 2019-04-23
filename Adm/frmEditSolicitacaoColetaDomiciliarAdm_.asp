<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>

<%
	Dim NumSolColeta
	Dim QtdCartuchos
	Dim QtdCartuchosRecebidos
	Dim DataSolicitacao
	Dim DataAprovacao
	Dim DataProgramada
	Dim DataEnvioTransp
	Dim DataReceb
	Dim StatusSol
	Dim MotivoStatus
	Dim Master
	Dim RazaoSocial
	Dim NomeFantasia
	Dim CEP
	Dim LogradouroColeta
	Dim NumEndColeta
	Dim CompEndColeta
	Dim MunEndColeta
	Dim UFEndColeta
	Dim DDDEndColeta
	Dim TelEndColeta
	Dim ContatoColeta
	Dim ReqColetaDomiciliar
	Dim NumRecTransportadora
	Dim IDCliente
	Dim IDTransp
	Dim StatusAprovar
	Dim StatusAtualizar
	Dim IdTranspHidden

    Dim Observacao

    Dim DadosCadastrais
    Dim DadosFaturamento
    Dim DadosColeta
	
	ReqColetaDomiciliar = Request.QueryString("iscoletadomiciliar")
	
	Sub GetSolicitacao()
		Dim sSql, arrSolicitacao, intSolicitacao, i
		sSql = "SELECT " & _ 
				"[Status_coleta_idStatus_coleta], " & _ 
				"[numero_solicitacao_coleta], " & _ 
				"[qtd_cartuchos], " & _ 
				"[qtd_cartuchos_recebidos], " & _ 
				"[data_solicitacao], " & _ 
				"[data_aprovacao], " & _ 
				"[data_programada], " & _ 
				"[data_envio_transportadora], " & _ 
				"[data_entrega_pontocoleta], " & _ 
				"[data_recebimento], " & _ 
				"[motivo_status], " & _ 
				"[isMaster]," & _
                "[observacao]" & _ 
				"FROM [marketingoki2].[dbo].[Solicitacao_coleta] " & _
				"WHERE [idSolicitacao_coleta] = " & Request.QueryString("idsolic")
				
'		Response.Write sSql
'		Response.End()						
		Call search(sSql, arrSolicitacao, intSolicitacao)

		If intSolicitacao > -1 Then
			For i=0 To intSolicitacao
				StatusSol 							= arrSolicitacao(0,i)
				NumSolColeta 						= arrSolicitacao(1,i)
				QtdCartuchos 						= arrSolicitacao(2,i)
				QtdCartuchosRecebidos 	= arrSolicitacao(3,i)
				DataSolicitacao 				= arrSolicitacao(4,i)
				DataAprovacao 		 			= arrSolicitacao(5,i)
				DataProgramada 					= arrSolicitacao(6,i)		
				DataEnvioTransp 				= arrSolicitacao(7,i)
				DataReceb 							= arrSolicitacao(9,i)
				MotivoStatus 						= arrSolicitacao(10,i)
                Observacao                      = arrSolicitacao(12,i)
			Next
			
		If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
			If Not isNull(DataSolicitacao)				Then DataSolicitacao	= FormatDateTime(DataSolicitacao, 2) 						
			If Not isNull(DataAprovacao)				Then DataAprovacao		= FormatDateTime(DataAprovacao, 2) 						
			If Not isNull(DataProgramada)				Then DataProgramada		= FormatDateTime(DataProgramada, 2)						
			If Not isNull(DataEnvioTransp)				Then DataEnvioTransp	= FormatDateTime(DataEnvioTransp, 2)						
			If Not isNull(DataReceb)					Then DataReceb			= FormatDateTime(DataReceb, 2)
		Else			
'			response.write DateRight(FormatDateTime(DataProgramada, 2))
			If Not isNull(DataSolicitacao)Then DataSolicitacao	= DateRight(FormatDateTime(DataSolicitacao, 2))
			If Not isNull(DataAprovacao)				Then DataAprovacao		= DateRight(FormatDateTime(DataAprovacao, 2))
			If Not isNull(DataProgramada)				Then DataProgramada		= DateRight(FormatDateTime(DataProgramada, 2))
			If Not isNull(DataEnvioTransp)				Then DataEnvioTransp	= DateRight(FormatDateTime(DataEnvioTransp, 2))
			If Not isNull(DataReceb)					Then DataReceb			= DateRight(FormatDateTime(DataReceb, 2))
		End If	

		End If
		Call GetCliente()
		Call GetEnderecoColeta()
		Call GetIDTransp()		
        Call GetDadosCadastrais()
        Call GetDadosFaturamento()
	End Sub
	
	function corrigiData(data)
		dim data1
		dim data_correta
		data1 = split(data, "/")
		if len(data1(0)) = 1 then
			data_correta =  "0"&data1(0)&"/"
		else
			data_correta =  data1(0)&"/"
		end if	
		if len(data1(1)) = 1 then
			data_correta =  data_correta & "0"&data1(1)&"/"
		else
			data_correta =  data_correta & data1(1)&"/"
		end if	
			data_correta =  data_correta & data1(2)
		corrigiData = data_correta
	end function
	
	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano
		
		Dia = day(sData)
		if Dia < 10 then Dia = "0" & Dia

		Mes = month(sData)
		if Mes < 10 then Mes = "0" & Mes

		Ano = year(sData)

		DateRight = Dia & "/" & Mes & "/" & Ano
	End Function
	
	Sub GetCliente()
		Dim sSql, arrCliente, intCliente, i
		sSql = "SELECT " & _ 
						"B.[razao_social], " & _
						"B.[nome_fantasia], " & _
						"B.[cnpj], " & _ 
						"B.[compl_endereco_coleta], " & _ 
						"B.[numero_endereco_coleta], " & _
						"B.[contato_respcoleta], " & _ 
						"B.[ddd_respcoleta], " & _ 
						"B.[telefone_respcoleta], " & _
						"B.[idClientes] " & _
						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
						"ON A.[Clientes_idClientes] = B.[idClientes] " & _
						"WHERE A.[Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
						
		Call search(sSql, arrCliente, intCliente)

		If intCliente > -1 Then
			For i=0 To inCliente
				RazaoSocial   = arrCliente(0,i)
				NomeFantasia  = arrCliente(1,i)
				NumEndColeta  = arrCliente(4,i)
				CompEndColeta = arrCliente(3,i)
				IDCliente	  = arrCliente(8,i) 
				DDDEndColeta  = arrCliente(6,i)
				TelEndColeta  = arrCliente(7,i)
				ContatoColeta = arrCliente(5,i) 	

			Next
		End If
	End Sub

	Sub GetEnderecoColeta()
		Dim sSql, arrEnd, intEnd, i
		
		sSql =  "SELECT " & _
				"A.[cep_coleta], " & _ 
				"rtrim(A.[logradouro_coleta]), " & _ 
				"rtrim(A.[bairro_coleta]), " & _ 
				"rtrim(A.[municipio_coleta]), " & _ 
				"A.[estado_coleta], " & _ 
				"A.[numero_endereco_coleta], " & _ 
				"A.[comp_endereco_coleta], " & _ 
				"A.[ddd_resp_coleta], " & _ 
				"A.[telefone_resp_coleta], " & _ 
				"A.[contato_coleta], " & _ 
                "A.[ramal_resp_coleta], " & _ 
                "A.[depto_resp_coleta] " & _ 
				"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _ 
				"WHERE Solicitacao_coleta_idSolicitacao_coleta = " & Request.QueryString("idsolic")
		
				'A.[cep_coleta]					= 0
				'A.[logradouro_coleta]			= 1  
				'A.[bairro_coleta]				= 2
				'A.[municipio_coleta]			= 3 
				'A.[estado_coleta]				= 4 
				'A.[numero_endereco_coleta]		= 5
				'A.[comp_endereco_coleta]       = 6
				'A.[ddd_resp_coleta]			= 7 
				'A.[telefone_resp_coleta]		= 8
				'A.[contato_coleta]				= 9

'		response.Write sSql
'		response.End						
						
		Call search(sSql, arrEnd, intEnd)
		If intEnd > -1 Then
			For i=0	To intEnd
				CEP 			 = arrEnd(0,i)
				LogradouroColeta = arrEnd(1,i)
                BairroColeta     = arrEnd(2,i)
				MunEndColeta 	 = arrEnd(3,i)
				UFEndColeta 	 = arrEnd(4,i)
				NumEndColeta  	 = arrEnd(5,i)
				CompEndColeta 	 = arrEnd(6,i)
				DDDEndColeta  	 = arrEnd(7,i)
				TelEndColeta	 = arrEnd(8,i)
				ContatoColeta 	 = arrEnd(9,i) 	
                ramalColeta      = arrEnd(10,i)
                deptoColeta      = arrEnd(11,i)

                CEP = Mid(CEP,1,5) & "-" & Mid(CEP,6,3)'12345000

                DadosColeta     = LogradouroColeta & "," & NumEndColeta & chr(13) & CompEndColeta & chr(13) & BairroColeta & chr(13) & CEP & chr(13) & MunEndColeta & " - " & UFEndColeta & chr(13) & _
                    "Contato: " & ContatoColeta & chr(13) & _
                    "("& DDDEndColeta & ")" & TelEndColeta & chr(13) & "Ramal: " & ramalColeta & chr(13) & _
                    "Depto: " & deptoColeta & chr(13) & _ 
                    "Observações:" & Observacao
                    
			Next					
		End If
	End Sub
	
	Sub GetStatusColeta()
		Dim sSql, arrStatus, intStatus, i
		Dim sSelected 
		sSelected = ""
		
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _ 
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta]"
		Call search(sSql, arrStatus, intStatus)						
		If intStatus > -1 Then
			For i=0 To intStatus
				If StatusSol = arrStatus(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value='"&arrStatus(0,i)&"' "&sSelected&">"&arrStatus(1,i)&"</option>"
			Next
		End If
	End Sub

	Sub GetDescStatusColeta(lId)
		Dim sSql, arrStatus, intStatus, i
		
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _ 
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta] " & _
						"WHERE idStatus_coleta = " & lId
		Call search(sSql, arrStatus, intStatus)						
		If intStatus > -1 Then
			Response.Write arrStatus(1,0)
		End If
	End Sub	
	
	Sub RequestForm()
		'response.write "status cbStatusSolColeta: " & left( Request.Form("cbStatusSolColeta"),1) & " </br>" 
		QtdCartuchosRecebidos 		= Request.Form("txtQtdCatuchosRecebidos")
		DataAprovacao 				= Request.Form("txtDataAprovacao")
		DataProgramada 				= Request.Form("txtDataProgramada")
		DataReceb 					= Request.Form("txtDataRecebimento")
		StatusSol 					= Request.Form("cbStatusSolColeta")
		MotivoStatus 				= Request.Form("txtMotivoStatus")
        Observacao                  = Request.Form("txtObservacao")
		DataEnvioTransp				= Request.Form("txtDataEnvioTransportadora")
		NumRecTransportadora  		= Request.Form("txtNumConhTransportadora")
		IDTransp 					= Request.Form("cbTransp")

		If QtdCartuchosRecebidos 	= "" Then QtdCartuchosRecebidos				= "NULL" 
		If DataReceb 				= "" Then DataReceb 						= "NULL" Else DataReceb = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If 
		If DataProgramada 			= "" Then DataProgramada 					= "NULL" Else DataProgramada = "CONVERT(DATETIME, '"&FormatDate(DataProgramada)&"')" End If 
		If DataAprovacao 			= "" Then DataAprovacao 					= "NULL" Else DataAprovacao = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If 
		If DataEnvioTransp 			= "" Then DataEnvioTransp 					= "NULL" Else DataEnvioTransp = "CONVERT(DATETIME, '"&FormatDate(DataEnvioTransp)&"')" End If
		If NumRecTransportadora  	= "" Then NumRecTransportadora				= "NULL" 
		If DataEntPontoColeta 	 	= "" Then DataEntPontoColeta 				= "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
		If IDTransp 				= "" Then IDTransp 							= "NULL" 
		If MotivoStatus 			= "" Then MotivoStatus 						= "NULL"
        If Observacao               = "" Then Observacao                        = "NULL"
		'response.write "fim resquestform </br>"
	End Sub
	
	Function GetIDTranspDefault()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT Transportadoras_idTransportadoras FROM Clientes WHERE idClientes = " & IDCliente
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If Not isNull(arrTransp(0,i)) Then
					GetIDTranspDefault = arrTransp(0,i)
				End If
			Next
		End If
	End Function 
	
	Sub GetTransp()
		Dim sSql, arrTransp, intTransp, i
		Dim sSelected
		sSql = "SELECT [idTransportadoras] " & _
			   ",[nome_fantasia] " & _
			   "FROM [marketingoki2].[dbo].[Transportadoras] " & _
			   "WHERE [status] = 1"
		
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If GetIDTransp() = arrTransp(0,i) Then
					sSelected = "selected"
				ElseIf GetIDTranspDefault() = arrTransp(0,i) Then
					sSelected = "selected"
				Else	
					sSelected = ""
				End If	
				Response.Write "<option value="&arrTransp(0,i)&" "&sSelected&">"&arrTransp(1,i)&"</option>"
			Next
		End If	   
	End Sub
	
	Sub SubmitForm()
        'response.Write Request.ServerVariables("HTTP_METHOD") 
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			Call UpdateSol()
		Else
			Call GetSolicitacao()
		End If

	End Sub

    function GetDadosFaturamento()
        dim sql, arr ,intarr, i, sSql
        ssql = "select top 1 cep, logradouro, rtrim(bairro), rtrim(municipio), estado,  compl_endereco, numero_endereco from [dbo].[cep_consulta_has_Clientes] as a " & _
	            "inner join Clientes as b on a.Clientes_idClientes = b.idClientes " & _
	            "inner join Solicitacao_coleta_has_clientes as c on a.Clientes_idClientes = c.Clientes_idClientes " & _
                "where Solicitacao_coleta_idSolicitacao_coleta = " & Request.QueryString("idsolic") & " " & _
                "and a.isEnderecoComum = 1 "
    
        Call search(sSql, arrEnd, intEnd)

		If intEnd > -1 Then
			For i=0	To intEnd
                CEP = arrEnd(0,i)
                CEP = Mid(CEP,1,5) & "-" & Mid(CEP,6,3)'12345000

                DadosFaturamento  = arrEnd(1,i) & "," & arrEnd(6,i) & chr(13) & arrEnd(5,i) & chr(13) & _
                    arrEnd(2,i) & chr(13) & _ 
                    CEP & chr(13) & arrEnd(3,i) & " - " & arrEnd(4,i)
            Next
        End if

    End Function


    function getdadoscadastrais()
        dim sql, arr ,intarr, i, sSql

        sSql = "select top 1 razao_social, nome_fantasia, cnpj, inscricao_estadual, ddd, telefone, compl_endereco, numero_endereco " & _
                "from Clientes inner join Solicitacao_coleta_has_clientes "& _
	            "on Clientes_idClientes = idClientes " & _
                "where Solicitacao_coleta_idSolicitacao_coleta = " & Request.QueryString("idsolic")
	    
        Call search(sSql, arrEnd, intEnd)

		If intEnd > -1 Then
			For i=0	To intEnd
                DadosCadastrais  = arrEnd(0,i) & chr(13) & arrEnd(1,i) & chr(13) & arrEnd(2,i) & chr(13) & arrEnd(3,i) & chr(13) & _ 
                "(" & arrEnd(4,i) & ")" & arrEnd(5,i)
            Next
        End if
    End function



	function getCheckCliente(id)
		dim sql, arr ,intarr, i
		
		sql = "SELECT A.[idClientes] " & _
			  ",A.[status_cliente] " & _
			  "FROM [marketingoki2].[dbo].[Clientes] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_clientes] AS B " & _
			  "ON B.[Clientes_idClientes] = A.[idClientes] " & _
			  "WHERE A.[status_cliente] = 1 AND B.[Solicitacao_coleta_idSolicitacao_coleta] = " & id
		
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			if not isnull(arr(0,i)) and not isempty(arr(0,i)) then
				getCheckCliente = true
			else
				getCheckCliente = false
			end if
		else
			getCheckCliente = false
		end if
	end function
	
	Sub UpdateSol()
		Dim sSql, arrSol, intSol, i
		
		'Response.Write ssql & "StatusSol " & StatusSol & "<hr>"
		
		if getCheckCliente(request.form("id")) then
		
			sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
					"SET [Status_coleta_idStatus_coleta]	= "&StatusSol&" " & _
					  ",[qtd_cartuchos_recebidos]			= "&QtdCartuchosRecebidos&" " & _
					  ",[data_aprovacao]					= "&DataAprovacao&" " & _
					  ",[data_programada]					= "&DataProgramada&" " & _
					  ",[data_envio_transportadora]			= "&DataEnvioTransp&" " & _
					  ",[data_entrega_pontocoleta]			= NULL " & _
					  ",[data_recebimento]					= "&DataReceb&" " & _
					  ",[motivo_status]						= '"&MotivoStatus&"' " & _
					"WHERE [idSolicitacao_coleta]			= " & Request.Form("id")
			'Response.Write ssql & "<hr>"
	'Response.End
			Call exec(sSql)
			If CInt(Request.Form("hiddenReqColetaDomiciliar")) = 1 Then
				sSql = "SELECT [Solicitacao_coleta_idSolicitacao_coleta] " & _
							  ",[Transportadoras_idTransportadoras] " & _
							  ",[numero_reconhecimento_transportadora] " & _
						  "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
						  "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.Form("id")
				'Response.Write ssql & " - 1111<hr>"
				Call search(sSql, arrSol, intSol)		  
				If  intSol > -1 Then
					If IDTransp <> "" And IDTRansp <> "NULL" And IDTRansp <> "-1" Then
						sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _ 
								 "SET [Transportadoras_idTransportadoras] = "&IDTransp&", " & _
								 "[numero_reconhecimento_transportadora] = '"&NumRecTransportadora&"' " & _
								 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.Form("id")
						'Response.Write sSql & " - foi"
						'Response.End()		 
						Call exec(sSql)					 
					End If
				Else
					'Response.Write IDTransp & " - teste"
					If IDTransp <> "" And IDTRansp <> "NULL" And IDTRansp <> "-1" Then
						sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
									   "([Solicitacao_coleta_idSolicitacao_coleta] " & _
									   ",[Transportadoras_idTransportadoras] " & _
									   ",[numero_reconhecimento_transportadora]) " & _
								 "VALUES " & _
									   "("&Request.Form("id")&" " & _
									   ","&IDTransp&" " & _
									   ",'"&NumRecTransportadora&"')"
						'Response.Write sSql
						'Response.End()		 
						Call exec(sSql)						   
					End If 
				End If
			End If			
		else
			response.write "<script>alert('O Cliente desta Solicitação não tem um STATUS para efetuar essa operação')</script>"
		end if	
		response.write "<script>window.opener.location.reload();</script>"
		Response.Write "<script>window.parent.close();</script>"							
	End Sub
	
	Function GetIDTransp()
		Dim sSql, arrId, intId, i
		Dim Ret
		'sSql = "SELECT [Transportadoras_idTransportadoras] " & _
		'			 ",[numero_reconhecimento_transportadora] " & _
		'	  	 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
		'			 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")

		sSql = "SELECT Transportadoras_idTransportadoras FROM Clientes WHERE idClientes = " & IDCliente			   

		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(0,i)
			Next
		End If
		GetIDTransp = Ret				 
	End Function
	
	Function isColetaEmail()
		Dim sSql, arrTransp, intTransp, i
		Dim IDTransp
		Dim Ret
		
		If GetIDTransp() <> "" Then
			IDTransp = GetIDTransp()
		ElseIf GetIDTranspDefault() <> "" Then 	
			IDTransp = GetIDTranspDefault()
		End If
		
		If IDTransp <> "" Then
			sSql = "select iscoletaemail from transportadoras where idtransportadoras = " & IDTransp
			
			Call search(sSql, arrTransp, intTransp)
			If intTransp > -1 Then
				For i = 0 To intTransp
					Ret = arrTransp(0,i)	
				Next
			Else
				Ret = 0		
			End If
		Else
			Ret = 0
		End If	
		
		isColetaEmail = Ret
	End Function
	
	Function GetNumRecTransportadora()
		Dim sSql, arrId, intId, i
		Dim Ret
		sSql = "SELECT [Transportadoras_idTransportadoras] " & _
				 ",[numero_reconhecimento_transportadora] " & _
			  	 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
				 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(1,i)
			Next
		End If
		GetNumRecTransportadora = Ret				 
	End Function
	
	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia

		Dia = Left(sDate, 2)
		Mes = Mid(sDate, 4, 2)
		Mes = Replace(Mes, "/" ,"")
		Ano = Right(sDate, 4)

		'Dia = day(sDate)
		'Mes = month(sDate)
		'Ano = year(sDate)
		
		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function
	
	Call SubmitForm()
	
	If CInt(StatusSol) = 1 Or CInt(StatusSol) = 3 Then
		StatusAprovar = true 
	Else
		StatusAprovar = false
	End if	
	
	If Not CInt(StatusSol) = 1 And Not CInt(StatusSol) = 3 and not cint(StatusSol) = 6 and not cint(StatusSol) = 4 Then
		StatusAtualizar = True
	Else
		StatusAtualizar = False
	End If
	
	Sub GetDescTransp()
		Dim sSql, arrTransp, intTransp, i
		Dim sSelected

		if len(trim(GetIDTransp())) then
			sSql = "SELECT [idTransportadoras] " & _
				   ",[nome_fantasia] " & _
				   "FROM [marketingoki2].[dbo].[Transportadoras] " & _
				   "WHERE [idTransportadoras] = " & GetIDTransp()

			Call search(sSql, arrTransp, intTransp)

			If intTransp > -1 Then
				IdTranspHidden = arrTransp(0,i)
				Response.Write arrTransp(1,i)
				
			End If	   
		end if
	End Sub
%>
<html>

<head>
    <link rel="stylesheet" type="text/css" href="../css/geral.css">
    <script language="javascript" src="js/frmEditSolicitacaoColetaDomiciliarAdmLc.js"></script>
    <script language="JavaScript" src="js/CalendarPopup.js"></script>
    <script language="JavaScript">

        var cal = new CalendarPopup();
    </script>
    <title><%=TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style type="text/css">
        .auto-style1 {
            width: 8%;
        }

        .auto-style4 {
        }

        .auto-style5 {
            font-size: 12px;
            text-align: left;
        }

        .auto-style6 {
            width: 177px;
        }

        #numsolcoleta {
            text-align: right;
        }

        .auto-style8 {
            font-size: 12px;
            text-align: right;
        }

        #tableEditSolicitacaoColetaAdm3 {
            width: 265px;
        }

        .auto-style9 {
            font-size: 12px;
            text-align: right;
            width: 25%;
            margin-left:20px;
        }
        .auto-style10 {
            font-size: 12px;
            font-weight:bold;
        }
        .auto-style11 {
            font-size: 1px;
            height: 1px;
        }
        .auto-style12 {
            height: 1px;
        }
    </style>
</head>

<body>
    <div id="container">
        <div id="conteudo" style="height: 100%;">
            <form action="frmEditSolicitacaoColetaDomiciliarAdm.asp" name="frmEditSolicitacaoColetaDomiciliarAdm" method="POST">
                <input type="hidden" name="hiddenReqColetaDomiciliar" value="<%=ReqColetaDomiciliar%>" />
                <input type="hidden" name="id" value="<%=Request.QueryString("idsolic")%>" />
                <input type="hidden" name="hiddenIsColetaEmail" value="<%=isColetaEmail()%>" />
                <input type="hidden" name="hiddenIdCliente" value="<%=IDCliente%>" />
                <table cellpadding="1" cellspacing="1" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
                    <tr>
                        <td id="explaintitle" colspan="6" align="center">Administrar Solicitação de Coleta</td>
                    </tr>
                    <tr id="trnumsolcoleta">
                        <td colspan="6">&nbsp;</td>
                    </tr>
                    <tr id="trnumsolcoleta">
                        <td align="right" class="auto-style9" style="text-align: right;">Solic.Coleta:</td>
                        <td colspan="1" class="auto-style10"><%=NumSolColeta%></td>
                        <td colspan="1" align="right" class="auto-style8" style="text-align: right;">Qtd.Cartuchos:</td>
                        <td colspan="1" class="auto-style10"><%=QtdCartuchos%></td>
                        <td colspan="1" class="auto-style8" style="text-align: right;">&nbsp;</td>
                        <td colspan="1">&nbsp;</td>
                    </tr>
                    <tr id="trnumsolcoleta">
                        <td class="auto-style11"></td>
                        <td colspan="1" class="auto-style12"></td>
                        <td colspan="1" class="auto-style12"></td>
                        <td colspan="1" class="auto-style12"></td>
                        <td colspan="1" class="auto-style12"></td>
                        <td colspan="1" class="auto-style12"></td>
                    </tr>
                    <tr id="trnumsolcoleta">
                        <td align="right" class="auto-style9" style="text-align: right;">Dt.Solic.do Pedido de Coleta:</td>
                        <td colspan="1" class="auto-style10"><%=DataSolicitacao%></td>
                        <td colspan="1">&nbsp;</td>
                        <td colspan="1">&nbsp;</td>
                        <td colspan="1">&nbsp;</td>
                        <td colspan="1">&nbsp;</td>
                    </tr>
                    <tr id="tridcliente">
                        <td colspan="5">&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr id="tridcliente">
                        <td colspan="5" style="font-size: 12px; background-color: #FF6A6A; color: #fff;">Dados Cadastrais do Cliente</td>
                        <td></td>
                    </tr>
                    <tr id="tridcliente">
                        <td colspan="6">
                            <textarea name="txtDadosCadastrais" style="width: 660px; height: 95px;" disabled><%=DadosCadastrais%></textarea></td>
                    </tr>
                    <tr id="tridcliente">
                        <td colspan="6">&nbsp;</td>
                    </tr>
                    <tr id="trcepcoleta">
                        <td colspan="3" style="font-size: 12px; background-color: #FF6A6A; color: #fff;">Dados para Faturamento:</td>
                        <td style="font-size: 12px; background-color: #FF6A6A; color: #fff;" colspan="2">Dados para Coleta:</td>
                    </tr>

                    <tr id="trcepcoleta">
                        <td colspan="3">
                            <textarea name="txtDadosFaturamento" style="width: 302px; height: 183px;" disabled><%=DadosFaturamento%></textarea></td>
                        <td align="right" class="auto-style5" colspan="2">
                            <textarea name="txtDadosColeta" style="width: 302px; height: 183px;" disabled><%=DadosColeta %></textarea></td>
                        <td>&nbsp;</td>
                        <td class="auto-style1">&nbsp;</td>
                    </tr>

                    <tr id="trcepcoleta">
                        <td colspan="3">&nbsp;</td>
                        <td align="right" class="auto-style5" colspan="2">&nbsp;</td>
                        <td>&nbsp;</td>
                        <td class="auto-style1">&nbsp;</td>
                    </tr>

                    <tr id="trqtdcartuchosrecebidos">
                        <td align="right" class="auto-style9">
                            <label id="qtdcartuchosrecebidos">Qtd. cartuchos recebidos: </label>
                        </td>
                        <td colspan="1">
                            <input type="text" <%If isNull(QtdCartuchosRecebidos) Then%>class="textreadonly" <%ELse%>class="textreadonly" <%End If%> name="txtQtdCatuchosRecebidos" readonly="readonly" value="<%=QtdCartuchosRecebidos%>" size="4" />&nbsp;<img src="img/produtos.gif" align="absmiddle" class="imgexpandeinfo" width="25" height="22" name="listprodutos" alt="Produtos" onclick="javascript:window.open('frmListaProdutosSolicitacao.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /></td>
                    </tr>
                    <tr>
                        <td align="right" class="auto-style9">
                            <label id="status">Status: </label>
                        </td>
					<td colspan="6">
						<select name="cbStatusSolColeta" class="select">
						<option value="-1"> --- Selecione --- </option>						
						<%Call GetStatusColeta()%>
						</select>
						<!--
						<call GetDescStatusColeta(StatusSol)%>
						<input type="hidden" id="cbStatusSolColeta" name="cbStatusSolColeta" value="<=StatusSol%>">-->
					</td>
                    </tr>
                    <tr>
                        <td align="right" class="auto-style9">
                            <label id="motivostatus">Motivo status:</label></td>
                        <td colspan="2">
                            <textarea name="txtMotivoStatus" style="width: 250px; height: 100px;"><%If Not MotivoStatus = "NULL" Then Response.Write MotivoStatus End If%></textarea></td>
                    </tr>

                    <%If StatusSol <> 1 Then%>
                    <tr id="trdataaprovacao">
                        <td align="right" class="auto-style9">
                            <label id="dataaprovacao">Dt.Aprovação de Coleta: </label>
                        </td>
                        <td colspan="6">
                            <input type="text" <%If isNull(DataAprovacao) Then%> class="textreadonly" value="<%=DataAprovacao%>" <%ELse%>class="textreadonly" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" readonly="readonly" size="13" maxlength="10" onkeypress="date(this)" /></td>
                    </tr>
                    <tr id="trdataenviotransportadora">
                        <td align="right" class="auto-style9">Dt.Coleta no Cliente:</td>
                        <td valign="bottom" colspan="6">
                            <input type="text" <%If isNull(DataEnvioTransp) Then%> class="textreadonly" <%ELse%> class="textreadonly" <%End If%> name="txtDataEnvioTransportadora" value="<%=DataEnvioTransp%>" size="13" maxlength="10" />
                            <a href="#" onclick="cal.select(document.forms['frmEditSolicitacaoColetaDomiciliarAdm'].txtDataEnvioTransportadora,'anchor1','dd/MM/yyyy'); return false;" name="anchor1" id="anchor1">
                                <img align="absmiddle" src="img/btn_calendario.gif" border="0"></a>
                        </td>
                    </tr>
                    <tr id="tridtransportadora">
                        <td align="right" class="auto-style9">Transportadora:</td>
                        <td colspan="6">

                            <%Call GetDescTransp()%>
                            <!--<input type="text" id="cbTransp" name="cbTransp" value="<%=IdTranspHidden%>">-->
                        </td>
                    </tr>
                    <tr id="trdataprogramada">
                        <td align="right" class="auto-style9">
                            <label id="dataprogramada">Dt.Programada p/ Coleta: </label>
                        </td>
                        <td valign="bottom" colspan="6">
                            <input type="text" <%If isNull(DataProgramada) Then%>class="textreadonly" <%ELse%>class="textreadonly" <%End If%> name="txtDataProgramada" value="<%=DataProgramada%>" size="13" maxlength="10" />
                            <a href="#" onclick="cal.select(document.forms['frmEditSolicitacaoColetaDomiciliarAdm'].txtDataProgramada,'anchor1','dd/MM/yyyy'); return false;" name="anchor1" id="anchor1">
                                <img align="absmiddle" src="img/btn_calendario.gif" border="0"></a>
                        </td>
                    </tr>
                    <tr id="trnumconhtransportadora">
                        <td align="right" class="auto-style9">
                            <label id="numconhtransportadora">N.Conhecimento:&nbsp; </label>
                        </td>
                        <td colspan="6">
                            <input type="text" <%If isNull(GetNumRecTransportadora()) Or GetNumRecTransportadora() = "" Or GetNumRecTransportadora() = "NULL" Then%>class="textreadonly" value="" <%ELse%>class="textreadonly" value="<%=GetNumRecTransportadora()%>" <%End If%> name="txtNumConhTransportadora" size="15" /></td>
                    </tr>
                    <tr id="trdatarecebimento">
                        <td align="right" class="auto-style9">
                            <label id="datarecebimento">Dt.Chegada no Armazém:</label></td>
                        <td colspan="6">
                            <input type="text" <%If isNull(DataReceb) Or DataReceb = "" Then%>class="textreadonly" <%ELse%>class="textreadonly" <%End If%> name="txtDataRecebimento" value="<%=DataReceb%>" size="13" maxlength="10" readonly="readonly" /></td>
                    </tr>
                    <%End If%>

                    <tr <%If StatusAprovar=false Then%> style="visibility: hidden;" <%else %> style="visibility: visible" <%End If %>>
                        <td style="text-align: right;" class="auto-style9">
                            <label id="datarecebimento0">Transportadora:</label></td>
                        <td colspan="2">
                            <select name="cbTransp0" class="select">
                                <option value="-1">[Selecione]</option>
                                <%Call GetTransp()%>
                            </select><img src="img/transportadoras.gif" class="imgexpandeinfo" width="25" height="25" align="absmiddle" alt="Buscar Transportadora" onclick="window.open('frmTransportadoraLc.asp?idcli=<%=IDCliente%>','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" /></td>
                        <td class="auto-style4" colspan="2">
                            &nbsp;</td>
                        <td class="auto-style6">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;</td>
                        <td class="auto-style1">&nbsp;</td>
                    </tr>
                    <% If getCheckCliente(Request.QueryString("idsolic")) Then %>

                    <tr <%If StatusAprovar=false Then%> style="visibility: hidden;" <%else %> style="visibility: visible" <%End If %>>
                        <td style="text-align: right;" class="auto-style9">&nbsp;</td>
                        <td colspan="2">&nbsp;</td>
                        <td class="auto-style4" colspan="2">&nbsp;</td>
                        <td class="auto-style6">&nbsp;</td>
                        <td class="auto-style1">&nbsp;</td>
                    </tr>
                    <tr id="btnprove" <% If StatusAprovar=false Then %>style="display:none;" <%End If%>>
                        <td align="right" class="auto-style8">&nbsp;</td>
                        <td align="left" colspan="5">
                            <input type="button" class="btnformMaior" name="btnAprovar" value="Aprovar" onclick="aprovar('<%= Request.QueryString("idsolic") %>', '<%=left(NumSolColeta,1)%>')" />
                            <input type="button" class="btnformMaior" name="btnReprovar" value="Rejeitar" onclick="reprovar('<%= Request.QueryString("idsolic") %>')" />
                            <input type="button" class="btnformMaior" name="btnReprovar1" value="Cancelar" onclick="cancelar('<%= Request.QueryString("idsolic") %>')" />
                        </td>
                    </tr>
                    <tr id="btnatualizar" <% If Not StatusAtualizar Then %>style="display:none;" <%End If%>>
                        <td align="center" style="visibility:hidden;">
                            _____________________________</td>
                        <td align="left" colspan="5">
                            <input type="button" class="btnformMaior" name="btnAtualizar" value="Atualizar" onclick="validateForm()" /></td>
                    </tr>
                    <% End If %>
                </table>
            </form>
        </div>
    </div>
</body>
</html>
<%Call close()
%>

