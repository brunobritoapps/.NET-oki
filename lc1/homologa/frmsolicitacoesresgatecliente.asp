<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>

<%
    
    Sub SubmitForm()
        
        Set oCommand = Server.CreateObject("ADODB.Command")

		If Request.ServerVariables("HTTP_METHOD") = "GET" Then
            If Request.QueryString("query") = "DELETE" Then
                if Request.QueryString("id") <> "" Then
                    if Request.QueryString("sr") <> "" Then
                        
                        Dim sSql, arrSR, intSR, i
                        '
                        '
                        'recalcula o saldos e salva saldo recalculado.
                        dim sql, arr, intarr
                        dim nVlrPontuacao
                        dim nBonus

                        sql = " select numero_solicitacao from Solicitacao_Resgate_has_Solicitacao_Composicao "
                        sql = sql & " where numero_resgate = '" & Request.QueryString("sr") & "' "
                        sql = sql & " and numero_solicitacao <> '' "
                        
                        call search(sql, arr, intarr)

                        If intarr > -1 Then
                            For i = 0 To intarr
                                dim iJ
                                '
                                '
                                'Faz a devolução do saldo para a solicitação de coleta.
                                'Localiza o foi baixado e retorna o saldo.
                                sql = " select pontuacao, pontuacao_atingir, saldo, numero_solicitacao from bonus_gerado_clientes where numero_solicitacao  = '" & arr(0,i) & "' "
                                
                                call search(sql, arrJ, intArrJ)

                                If intArrJ > -1 Then
                                
                                    For iJ = 0 To intArrJ
                                        nVlrPontuacao   = nVlrPontuacao + arrJ(0,iJ)
                                        nBonus          = nBonus + arrJ(0,iJ)
                                    Next

                                End If

                                '
                                'depois de ler todo o bonus gerado. salva novamente o bonus.
                                sql = "update bonus_gerado_clientes set saldo = " & nVlrPontuacao & " where numero_solicitacao = '"& arr(0,i) & "'"
                                call exec(sql)

                                '
                                'zera novamente as variáveis.
                                nVlrPontuacao = 0

                            Next

                        End If

                        '
                        'após retornar saldo executa as deleções.

                        '
                        '
                        'select * from Solicitacao_Resgate_has_Solicitacao_Composicao where numero_resgate = 'R05140007084'
                        sSql = "delete from Solicitacao_Resgate_has_Solicitacao_Composicao where numero_resgate = '" & Request.QueryString("sr") & "' "
					    call exec(sSql)

                        sSql = "delete from Solicitacoes_resgate_Clientes where numero_solicitacao_geracao = '" & Request.QueryString("sr") & "' "
					    call exec(sSql)
    
                        'select * from solicitacao_coleta where idSolicitacao_coleta = 5984
                        sSql = "delete from solicitacao_coleta where idSolicitacao_coleta  = " & Request.QueryString("id")
                        call exec(sSql)

                        response.write "<script>alert('Resgate Excluído com sucesso! Foi retornado um Bônus de: '"& nBonus &"');</script>"

                    End if
                End if
            End if
        End if
        
    End Sub

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

    Function DateLc(sData)
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
		DateLc = Dia & "/" & Mes & "/" & Ano
	End Function


	function getSolicitacoesResgate()
		dim sql, arr, intarr, i
		dim arr2, intarr2, j
		dim html, style
		'sql = "SELECT distinct([numero_solicitacao_geracao]) " & _
		'	  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
		'	  "WHERE [idcliente] = " & session("IDCliente")
		sql = "select " & _
				"distinct(a.numero_solicitacao_geracao), " & _
				"a.idsolicitacao, " & _
				"b.qtd_cartuchos, " & _
				"b.data_recebimento, " & _
				"b.ismaster, " & _
				"d.typecolect, " & _
				"a.idcliente, " & _
				"b.status_coleta_idstatus_coleta,  " & _
				"e.status_coleta, " & _
				"a.data_solicitacao_resgate " & _
				"from solicitacoes_resgate_clientes as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.idsolicitacao = b.idsolicitacao_coleta " & _
				"left join clientes as d " & _
				"on a.idcliente = d.idclientes " & _
				"left join status_coleta as e " & _
				"on b.status_coleta_idstatus_coleta = e.idstatus_coleta where d.[idclientes] = " & session("IDCliente") & " " & _
                "order by data_solicitacao_resgate"
		'call search(sql, arr2, intarr2)
		'if intarr2 > -1 then
		'	for j=0 to intarr2
				'sql = "SELECT [numero_solicitacao_geracao] " & _
				'		  ",[idSolicitacoes_resgate] " & _
				'		  ",[data_solicitacao_resgate] " & _
				'		  ",[idcliente] " & _
				'		  ",[quantidade] " & _
				'		  ",[idsolicitacao] " & _
				'	  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
				'	  "WHERE [numero_solicitacao_geracao] = '" & arr2(0,j)&"'"
				'response.write sql
				'response.write sql&"<br />"
				call search(sql, arr, intarr)
				if intarr > -1 then
					for i=0 to intarr
						if i mod 2 = 0 then
							style="class=""classColorRelPar"" style=""cursor:pointer;"""
                            cor="class='classColorRelPar'"
						else
							style="class=""classColorRelImpar"" style=""cursor:pointer;"""
                            cor="class='classColorRelImpar'"
						end if

                        if arr(7,i) <> 1 Then
                            cstylebutton = "disabled"
                        else
                            cstylebutton = ""
                        End if

						html = html & "<tr>"
                        '
                        'peterson aquino 17-5-2014 id:8 - permitir excluir solicitações de resgate ainda não atendidas;
                        html = html & "<td "&cor& " style='width:5px;'><input type='submit' value='Excluir' onClick='lcDelSR(" & arr(1,i) & ","""&arr(0,i)&""")' name='query' id='"&arr(1,i)&"' " & " " & cstylebutton & "/></td>"
						html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(1,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(0,i)&"</td>"
						if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
							html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(1,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&DateRight(arr(9,i))&"</td>"
						else
							html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(1,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&DateLc(arr(9,i))&"</td>"
						end if
						html = html & "<td "&style&" onClick=""javascript:window.open('frmViewSolresgate.asp?idsol="&arr(1,i)&"','','width=720,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"">"&arr(2,i)&"</td>"
						html = html & "</tr>"
					next
				else
					html = html & "<tr>"
					html = html & "<td colspan=""3"" "&style&"><b>Nenhum registro encontrado</b></td>"
					html = html & "</tr>"
				end if
			'next
		'end if
		getSolicitacoesResgate = html
	end function

	Sub getStatusColeta()
		Dim sSql, arrStatus, intStatus, i
		Dim sSelected
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta]"
		Call search(sSql, arrStatus, intStatus)
		With Response
			If intStatus > -1 Then
				For i=0 To intStatus
					If Request.QueryString("StatusSol") = CStr(arrStatus(0,i)) Then
						sSelected = "selected"
					Else
						sSelected = ""
					End If
					.Write "<option value='"&arrStatus(0,i)&"' "&sSelected&">"&arrStatus(1,i)&"</option>"
				Next
			End If
		End With
	End Sub

    Call SubmitForm()
%>

<html>
<head>
<script language="javascript" type="text/javascript" src="js/frmSRClc.js"></script>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="0" width="100%">
						<tr>
							<td colspan="3" id="explaintitle" align="center">Painel de Controle</td>
						</tr>
						  <% If not isnull(session("IDCliente")) and not isempty(session("IDCliente")) Then %>
								<tr>
									<td id="explaintitle" align="right" colspan="3" style="padding:4px 4px 4px 4px;">
										<a href="frmoperacionalcliente.asp?logoff=true" style="color: #FFFFFF;">Logoff do Sistema</a>
									</td>
								</tr>
						  <% End If %>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="3" id="explaintitle" align="center">Solicitações de Resgate</td>
						</tr>
						<tr>
							<td colspan="3">
								<%if Session("isColetaDomiciliar") = 1 and Session("isMaster") = 1 and session("cod_cli_consolidador") = 0 and session("cod_bonus") <> "" then%>
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelSolPendente">
										<tr>
                                            <th>Operação</th>
											<th>N° Solicitação</th>
											<th>DT. Solic. Resgate</th>
											<th>Quantidade</th>
										</tr>
										<%=getSolicitacoesResgate()%>
									</table>
								<%end if%>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</table>
        <!--#include file="inc/i_bottom.asp" -->
        
    </div>
</div>
</body>
</html>
<%Call close()%>
