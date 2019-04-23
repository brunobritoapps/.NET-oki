<!--#include file="../_config/_config.asp" -->
<%
'|--------------------------------------------------------------------
'| Arquivo: frmAddContato.asp																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação : 15/04/2007																		 
'| Descrição: Arquivo de Formulário para cadastro de Contato
'|--------------------------------------------------------------------
%>
<%Call open()%>
<% Response.Charset="ISO-8859-1" %>
<%	
	Dim method
	Dim id
	
	method = Request.QueryString("sub")
	id = Request.QueryString("id")
	
	If method = "getcheckusercontato" Then
		Call getCheckUserContato(Request.QueryString("user"), Request.QueryString("senha"))
	End If
	
	Sub getCheckUserContato(User, Senha)
		Dim sSql, arrUser, intUser
		Dim Ret

		sSql = "SELECT " & _ 
						"[idContatos] " & _ 
						"FROM [marketingoki2].[dbo].[Contatos] " & _
						"WHERE [usuario] = '"&User&"' " & _ 
						"AND [senha] = '"&Senha&"'"

		Call search(sSql, arrUser, intUser)

		If intUser > -1 Then
			Response.Write "true"
		Else
			Response.Write "false"
		End If
	End Sub
%>
<%Call close()%>