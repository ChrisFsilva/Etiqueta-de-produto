<%@ Language="VBScript" CodePage="65001" %>

<%
Response.CodePage = 65001 ' Define o código de página para UTF-8
Response.CharSet = "UTF-8" ' Define o charset da resposta
%>

<%
Session.LCID = 1046   ' Português (Brasil)
SetLocale(1046)
%>

<%

Session.LCID = 1046
	
	Dim query, connQrcod, strConnQrcod
	Dim vlrStarMinimo, regexValorMinimo
	Set regexValorMinimo = New RegExp
	regexValorMinimo.Pattern = "^[0-9]+([.][0-9]{1,4})?$"
	regexValorMinimo.Global = False
	Set connQrcod = Server.CreateObject("ADODB.Connection")
	
	strConnQrcod = "Driver={MySQL ODBC 26.7 Unicode Driver};Server=caminho_server;Database=nome_banco;User=user_banco;Password=senha_banco;Option=3;"
	
	connQrcod.Open strConnQrcod
	
	query = ""
	'response.write "update</br>" 
	
	For i = 1 To Request.Form("etiqueta").Count

		id = Request.Form("etiqueta")(i)
		selecao = Request.Form("selecao[]")(i)
		vlr = Request.Form("vlrStar[]")(i)
		vlrComplemento = Request.Form("vlrComplemento[]")(i)
		vlrStarMinimo = Trim(Request.Form("vlrStarMinimo[]")(i))
		
		if selecao = 1 then
			if vlr > 0 then				
				 
				If Not regexValorMinimo.Test(vlrStarMinimo) Then
                    erroMsg = "Valor StarSoft Item Mínimo inválido na linha " & i & "."
                    Exit For
                End If

                query = "Update D41 SET VALOR = " & vlr & ", VALORCOMPLEMENTO = " & vlrComplemento & ", VALORMINIMO = " & vlrStarMinimo & " WHERE D41.D41_001_C = '" & id & "'" 
				
				response.write query
				
				On Error Resume Next
				
				connQrcod.Execute(query)
				
				If Err.Number <> 0 Then
					 erroMsg = "Erro na execução da query:<br>" & query & "<br><br>" & _
                           "Erro: " & Err.Number & " - " & Err.Description
					Exit For   ' <<< sai do FOR
				End If		
					
				On Error GoTo 0
					
			end if
		end if
	
	Next	
	
	If erroMsg <> "" Then
		Response.Write erroMsg
		Response.End
	Else
		Response.Write "<script>alert('Atualizado com sucesso!'); window.location='index.asp';</script>"
		Response.End
	End If

	Response.Redirect "index.asp?msg=ok"
	
	connQrcod.close
	Set connQrcod = Nothing
	
%>