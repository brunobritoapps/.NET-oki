<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%
    
    Sub SubmitForm()
    
		Set oCommand = Server.CreateObject("ADODB.Command")
    
        response.write cCommand
    	response.write Request.ServerVariables("HTTP_METHOD") 
        response.Write Request.QueryString("query")

		If Request.ServerVariables("HTTP_METHOD") = "GET" Then
            If Request.QueryString("query") = "DELETE" Then
                '
                'pega o numero do coisa
                 cNumSR = Request.QueryString("id")

                'faz a exclusão do registro 
	            sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
			         "SET [Status_coleta_idStatus_coleta] = 2 " & _
				    ",[data_aprovacao] = GetDate() " & _
				    "WHERE [idSolicitacao_coleta] = " & Request.QueryString("id")

            End If
            
        End If        
    
        response.write cNumSR

    End Sub


	Call SubmitForm()
%>