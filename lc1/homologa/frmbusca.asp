<!--#include file="_config/_config.asp" -->
<!--#include file="inc/i_banner.asp" -->
<%Call open()%>
<%
	function getBusca()
		dim sql, arr, intarr, i
		dim arr2, intarr2, i2
		dim arr3, intarr3, i3
		dim soma
		dim html
		sql = "SELECT [idtexto] " & _
					  ",[texto] " & _
					  ",[area] " & _
				  "FROM [marketingoki2].[dbo].[Home_Textos] WHERE [texto] LIKE '%"&request.form("txtBusca")&"%'"
		call search(sql, arr, intarr)
		sql = "SELECT [idnoticia] " & _
				  ",[titulo] " & _
				  ",[text] " & _
				  ",[data] " & _
				  ",[fonte] " & _
				  ",[ativo] " & _
			  "FROM [marketingoki2].[dbo].[Home_Noticias] WHERE [text] LIKE '%"&request.form("txtBusca")&"%' "
		call search(sql, arr2, intarr2)	
		sql = "SELECT [data_inicio] " & _
					  ",[data_termino] " & _
					  ",[link] " & _
					  ",[busca] " & _
					  ",[imagem] " & _
					  ",[tipo] " & _
					  ",[idbanner] " & _
				  "FROM [marketingoki2].[dbo].[Home_Banners] WHERE [busca] LIKE '%"&request.form("txtBusca")&"%'"
		call search(sql, arr3, intarr3)
		if cint(intarr) > -1 then
			soma = soma + cint(intarr)	
		end if 
		if cint(intarr2) > -1 then
			soma = soma + cint(intarr2)	
		end if 
		if cint(intarr3) > -1 then
			soma = soma + cint(intarr3)	
		end if 
		soma = soma + 1
		if cint(soma) > -1 then
			html = html & "<tr>"
			html = html & "<td align=""center""><b style=""font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif"">Foram encontrado(s) "&soma&" conteúdo(s) com essa busca</b></td>"
			html = html & "</tr>"		    
			if intarr > -1 then
				for i=0 to intarr
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<div style=""overflow:hidden;width:510px;"">"
					html = html & "<table cellpadding=""1"" cellspacing=""1"" id=""tableRelCategories"">"
					html = html & "<tr>"
					html = html & "<td><b>"&arr(0,i)&":</b></td>"
					html = html & "</tr>"
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<a href=""index.asp?area="&arr(2,i)&""" class=""linkOperacional"">" & arr(1,i)
					html = html & "</a>"
					html = html & "</td>"
					html = html & "</tr>"
					html = html & "</table>"
					html = html & "</div>"
					html = html & "</td>"
					html = html & "</tr>"
				next
			end if		
			if intarr2 > -1 then
				for i2=0 to intarr2
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<div style=""overflow:hidden;width:510px;"">"
					html = html & "<table cellpadding=""1"" cellspacing=""1"" id=""tableRelCategories"">"
					html = html & "<tr>"
					html = html & "<td><b>"&arr2(1,i2)&":</b></td>"
					html = html & "</tr>"
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<a href=""index.asp?idnoticia="&arr2(0,i2)&""" class=""linkOperacional"">" & arr2(2,i2)
					html = html & "</a>"
					html = html & "</td>"
					html = html & "</tr>"
					html = html & "</table>"
					html = html & "</div>"
					html = html & "</td>"
					html = html & "</tr>"
				next
			end if
			if intarr3 > -1 then
				for i3=0 to intarr3
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<div style=""overflow:hidden;width:510px;"">"
					html = html & "<table cellpadding=""1"" cellspacing=""1"" id=""tableRelCategories"">"
					html = html & "<tr>"
					html = html & "<td><b>"&arr3(3,i3)&":</b></td>"
					html = html & "</tr>"
					html = html & "<tr>"
					html = html & "<td>"
					html = html & "<a href="""&arr3(2,i3)&""" class=""linkOperacional"">" & arr3(3,i3)
					html = html & "</a>"
					html = html & "</td>"
					html = html & "</tr>"
					html = html & "</table>"
					html = html & "</div>"
					html = html & "</td>"
					html = html & "</tr>"
				next
			end if
		else
			html = html & "<tr>"
			html = html & "<td align=""center""><b style=""font-size:10px;font-family:Verdana, Arial, Helvetica, sans-serif"">Não foi encontrado conteúdo para esta busca.</b></td>"
			html = html & "</tr>"		    
		end if	
		getBusca = html		
	end function
%>
<html>
<head>
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
        <td align="center" id="conteudo"><table width="750" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="510" valign="top"><p align="center">
			  	<table cellpadding="1" cellspacing="1" width="100%">
					<tr>
						<td id="explaintitle" align="center">Resultado da Busca</td>
					</tr>
					<%= getBusca() %>
				</table>
			  </p></td>
              <td width="240" valign="top"><div align="center"><br>
                  <%=getLogin()%> 
                 </div>
                <table width="200" border="0" align="center" cellpadding="1" cellspacing="1">
                  <%=getBannersLateral()%>
                </table>
                <p align="center"><br>
                </p></td>
            </tr>
          </table></td>
        <td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
      </tr>
    </table>
  </div>
  <!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>