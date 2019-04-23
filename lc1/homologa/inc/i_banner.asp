<%
	function getBannersLateral()
		dim sql, arr, intarr, i
		dim html
		html = ""
		sql = "SELECT [data_inicio] " & _
				  ",[data_termino] " & _
				  ",[link] " & _
				  ",[busca] " & _
				  ",[imagem] " & _
				  ",[tipo] " & _
				  ",[idbanner] " & _
			    "FROM [marketingoki2].[dbo].[Home_Banners] " & _
				"where data_inicio <= getdate() and data_termino >= dateadd(day, datediff(day, data_inicio, data_termino), data_inicio) and tipo = 'lateral'"
		call search(sql, arr, intarr)		
		if intarr > -1 then
			while i < intarr + 1 
				html = html & "<tr>"
				if left(request.servervariables("REMOTE_ADDR"), 3) = "127" then
					html = html & "<td><img src=""../sustentabilidade/adm/home/"&arr(4,i)&""" width=""220"" class=""imgexpandeinfo"" alt="""&arr(2,i)&""" onclick=""window.location.href='"&arr(2,i)&"'"" /></td>"
				else
					html = html & "<td><img src=""../sustentabilidade/adm/home/"&arr(4,i)&""" width=""220"" class=""imgexpandeinfo"" alt="""&arr(2,i)&""" onclick=""window.location.href='"&arr(2,i)&"'"" /></td>"
				end if	
				html = html & "</tr>"
				i = i + 1
			wend
		end if
		getBannersLateral = html
	end function
	
	function getLogin()
		dim html
		html = ""
		html = html & "<table width=""225"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		html = html & "<form action=""frmLoginCliente.asp"" name=""frmLoginCliente"" method=""POST"">"
		html = html & "<tr>"
		html = html & "<td><img src=""img/Box_Cadastro_top.gif"" width=""225"" height=""29""></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td background=""img/Bg_meio_box.gif"">" 
		html = html & "<table width=""210"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""2"" class=""textoHomeAzul"">"
		html = html & "<tr>"
		html = html & "<td>Entre com seu Login e senha para ter acesso ao portal " 
		html = html & "de devolu&ccedil;&atilde;o de suprimentos.</td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td height=""10""><img src=""img/_spacer.gif"" width=""1"" height=""5""></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td>Login:</td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td>"
		html = html & "<input name=""txtLogin"" type=""text"" class=""TextBox""></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td>Senha:</td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td>"
		html = html & "<input name=""txtSenha"" type=""password"" class=""TextBox""></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td><br><class=""textoHomeAzul""> Se não for cadastrado. <a href=""frmcadcliente.asp"" class=""linkOperacional"">clique aqui</a></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td><br><class=""textoHomeAzul""> Recuperar Senha <a href=""frmesqueci.aspx"" class=""linkOperacional"">clique aqui</a></td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td><img src=""img/_spacer.gif"" width=""1"" height=""5""></td>"
		html = html & "</tr>"
		html = html & "</table>"
		html = html & "</td>"
		html = html & "</tr>"
		html = html & "<tr>"
		html = html & "<td>"
		html = html & "<table width=""225"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		html = html & "<tr>" 
		html = html & "<td><img src=""img/corner_esq_box.gif"" width=""159"" height=""19""></td>"
		html = html & "<td><img src=""img/botao_enviar_box.gif"" width=""47"" height=""19"" class=""imgexpandeinfo"" alt=""Logar"" onclick=""document.frmLoginCliente.submit()""></td>"
		html = html & "<td><img src=""img/corner_dir_box.gif"" width=""19"" height=""19""></td>"
		html = html & "</tr>"
		html = html & "</table></td>"
		html = html & "</tr>"
		html = html & "</form>"
		html = html & "</table>"
		getLogin = html
	end function
	
	function getBannersRodape()
		dim sql, arr, intarr, i
		dim html
		html = ""
		sql = "SELECT [data_inicio] " & _
				  ",[data_termino] " & _
				  ",[link] " & _
				  ",[busca] " & _
				  ",[imagem] " & _
				  ",[tipo] " & _
				  ",[idbanner] " & _
			    "FROM [marketingoki2].[dbo].[Home_Banners] " & _
				"where data_inicio <= getdate() and data_termino >= dateadd(day, datediff(day, data_inicio, data_termino), data_inicio) and tipo = 'rodape'"
		call search(sql, arr, intarr)		
		if intarr > -1 then
			for i=0 to 0
				html = html & "<tr>"
				if left(request.servervariables("REMOTE_ADDR"), 3) = "127" then
					html = html & "<td><img src=""../sustentabilidade/adm/home/"&arr(4,i)&""" width=""468"" onclick=""window.location.href='"&arr(2,i)&"'"" /></td>"
				else
					html = html & "<td><img src=""../sustentabilidade/adm/home/"&arr(4,i)&""" width=""468"" onclick=""window.location.href='"&arr(2,i)&"'"" /></td>"
				end if	
				html = html & "</tr>"
			next
		end if
		getBannersRodape = html
	end function
%>
