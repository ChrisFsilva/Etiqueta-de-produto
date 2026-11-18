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
	 Dim enable 
	 enable = False	
	 
	Dim strConnIntra, strConnQrcod, strConnStar
	Dim connIntra, ConnStar, connQrcod
	Dim search, etiqueta, fornecedor, produto, queryMySql, CODFORNECEDOR, rs, rsStar, rsIntranet, strSQL1
	Dim vlrQrCode, vlrStar, contador
	Dim queryStarMinimo, rsStarMinimo, vlrStarMinimo, referenciaMinimo, descricaoMinimo
	
	vlrQrCode = 0 
	vlrStar = 0
	querySqlServer = ""
	queryMySql = ""
	etiqueta = trim(Request("etiqueta"))
	fornecedor = trim(Request("fornecedor"))
	produto = trim(Request("produto"))	
	contador = 0
	
	strConnIntra = "Provider=SQLOLEDB.1;Data Source=endereço_ERP;Initial Catalog=banco_ERP;User ID=readonlyuser;Password=Senha_ERP;"
	' strConnQrcod = "Driver={MySQL ODBC 9.6 Unicode Driver};Server=info_banco.mysql.dbaas.com.br;Database=info_banco;User=info_banco;password_banco;"
	strConnQrcod = "Driver={MySQL ODBC 26.7 Unicode Driver};Server=info_banco.mysql.dbaas.com.br;Database=info_banco;User=info_banco;password_banco;"
	strConnStar = "Provider=SQLOLEDB.1;Data Source=endereço_ERP;Initial Catalog=banco_ERP;User ID=readonlyuser;Password=Senha_ERP;"
	  
	Set connIntra = Server.CreateObject("ADODB.Connection")
	Set ConnStar = Server.CreateObject("ADODB.Connection")
	Set connQrcod = Server.CreateObject("ADODB.Connection")
	
	connIntra.Open strConnIntra
	connQrcod.Open strConnQrcod
	ConnStar.Open strConnStar

	strSQL1 = "SELECT MENU.*, USERS.USRCODE FROM USERMENU INNER JOIN MENU ON (USERMENU.MENUID = MENU.ID) INNER JOIN USERS ON  (USERMENU.USRID = USERS.ID) WHERE USERS.usrCode = '" & Session("USR_CODE") & "'"

	Set rsMenu = connIntra.Execute(strSQL1)				
		
	pos =  InStr(Session("name")," ")
	user = Session("name")
	
	if pos > 0 then
		user = Left(user, pos - 1)		
	end if
	
	If Not rsMenu.EOF Then		
		count = 1			
		Session("user") = rsMenu("usrcode")			
		Do Until rsMenu.EOF							
			if rsMenu("window") = 0 then
				menu = menu & "<li><a href='../" & rsMenu("pagina") & "'><button style='width:170px; background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;'>" & rsMenu("access") & "</button></a></li>"
			else								
				menu = menu & "<li><a onclick=""javascript:window.open(this.href, 'pop', 'width=2600, height=2300,scrollbars=yes' ); return false"" href=""" & rsMenu("pagina") & "?usrCode=" & rsMenu("usrcode") & "?name=" & user & """ target=""_blank""><button style='width:170px; background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;'>" & rsMenu("access") & "</button></a></li>"
				count = count +1
			end if				
			rsMenu.MoveNext				
		Loop
		
	end if
	
%>

	
<!DOCTYPE html>
<html lang="pt-BR">
<meta charset="UTF-8">
<head>
	<title>Brentwood - Etiqueta</title>	
	<style>
		* {
			box-sizing: border-box;
		}
		body {
			margin: 0;
			padding: 0;
			font-family: Arial, sans-serif;
			font-size: 16px;
			line-height: 1.5;
		}
		.row {
			display: flex;
			flex-wrap: wrap;
			margin: 0 -15px;
		}
		.col-menu {
			flex: 0 0 15%;
			max-width: 250px;
			padding: 0 15px;
		}
		.col-content {
			flex: 0 0 78%;
			max-width: 78%;
			padding: 0 15px;
		}
		.menu {
			padding: 5px;
			border-radius: 5px;
			box-shadow: 0 0 5px rgba(0,0,0,0.2);
			margin: 0 0 0 1em;
		}
		
		.menutop {
			padding: 20px;
			border-radius: 5px;
			box-shadow: 0 0 50px rgba(0,0,0,0.2);
			margin: 0 2.5em 0 1em;
		}
		.menu ul {
			list-style: none;
			padding: 0;
			margin: 0;
		}
		.menu li {
			margin-bottom: 10px;
		}
		.menu a {
			display: block;
			padding: 10px;
			border-radius: 5px;
			color: #333;
			text-decoration: none;
			transition: background-color 0.3s ease;
		}
		.menu a:hover {
			background-color: #069cc2;
			color: #fff;
		}
		.content {
			background-color: #fff;
			padding: 20px;
			border-radius: 5px;
			box-shadow: 0 0 5px rgba(0,0,0,0.2);
		}
		h1 {
			font-size: 24px;
			margin-top: 0;
			margin-bottom: 0px;
		}
		table {
		  border-collapse: collapse;
		  border-spacing: 0;
		  border: 0px solid #ddd;
		}

		th, td {
		  text-align: left;
		  padding: 16px;
		   font-size: 11px;
		}

		.tdcenter {
			text-align: center;
		}
		
		tr:nth-child(2n+1) {
		  background-color: #f2f2f2;
		}
		
		.field_set {
			border: 1px #D8D8D8 solid;
		}
		
		.valorDesconto {
			 width: 4%;
		}
		
		.transparent-button {
			background: none;
			border: none;
			color: blue; /* Cor do texto do bot&atilde;o */
			text-decoration: underline; /* Adicionando sublinhado para indicar que é um link */
			cursor: pointer; /* Alterando o cursor para indicar que é clicável */
			font-size: 16px;
			font-family: Arial, sans-serif;
		}
		/* Removendo o sublinhado quando o mouse passa sobre o bot&atilde;o */
			.transparent-button:hover {
			text-decoration: none;
		}
		.input-valor {
			width: 90px;
			padding: 4px;
			font-size: 11px;
			text-align: right;
		}

		.input-percentual {
			width: 70px;
			padding: 4px;
			font-size: 11px;
			text-align: right;
		}
	</style>
	
	<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.min.js"></script>

</head>
<body>

<script>
  $(document).ready(function() {
    // Ativa o datepicker para o campo de data com o ID "data"
    $("#data").datepicker({
      dateFormat: "dd/mm/yy", // Formato da data exibida e selecionada
      changeMonth: true, // Permite selecionar o mês a partir de um menu dropdown
      changeYear: true, // Permite selecionar o ano a partir de um menu dropdown
      yearRange: "1900:2050" // Define o intervalo de anos disponíveis
    });
  });
  
  function habilitarCampo(select) {
    var linha = select.closest('tr');
    var input = linha.querySelector("input[type='number']");

      input.readOnly = (select.value !== "1");
}
</script>

	<div class="menutop">
	<h1><font color="#9d46a3">Brentwood - <% response.write  Session("name") & "</br>" %>Atualizaç&atilde;o de Preços do QrCode com Base no StarSoft</font></br>
	</div><br>
	
	<div class="row">	
		<div class="col-menu">		
			<div class="menu">			
				<ul>
					<li><a href="#"><button style="width:170px; background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;">Menu Principal</button></a></li>
					<% response.write menu %>
					</br><li><a href="#"><button onclick="window.location='../logout.asp';" style="width:170px; background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;"><font color="red">Logout</font></button></a></li>
				</ul>		
			</div>	
		</div>	
		<div class="col-content">			
			<div class="content">
				<form action="" method="post">
					<label for="qrcode"><font color="#9d46a3">Etiqueta:</font></label>
					<input type="text" id="qrcode" name="etiqueta" size="10" autofocus>
					<label for="qrcode"><font color="#9d46a3">Cód. Fornecedor:</font></label>
					<input type="text" id="fornecedor" name="fornecedor" size="10" value="" autofocus>
					<label for="qrcode"><font color="#9d46a3">Produto:</font></label>
					<input type="text" id="produto" name="produto" size="10" autofocus>	
					<button type="submit">Pesquisar</button>
				</form>
			</div>  		
			</br>

	<%
		
	if fornecedor <> "" then			
		
		queryMySql = "SELECT * FROM D41 WHERE D41_001_C LIKE '%" & etiqueta & "%' and A08_001_C LIKE '%" & fornecedor & "%' and D04_008_C like '%" & produto & "%' and CONJUNTO <> 1 AND D41_001_C <> '' AND ATIVO = 1 order by A08_002_C, A08_001_C, D41_001_C"
		Set rs = connQrcod.Execute(queryMySql)
	
		count = 0
		CODFORNECEDOR = ""
				
		If Not (rs.EOF And rs.BOF) Then
		
		rs.MoveFirst
		
			if CODFORNECEDOR = "" then

				 response.write "<form method='post' style='' action='update3.asp'>"
				 response.write "<fieldset class='field_set'>"
				 response.write "<legend><font size='2px' color='#9d46a3'>Rela&ccedil;&atilde;o de Produtos: " & rs("A08_001_C") & " - " & rs("A08_002_C") & "</font></legend>"			
				 response.write "<table width='100%'><tr><th width='8%'>Etiqueta</th><th>Produto</th><th>Atualizar</th><th class='input-valor'>Valor QrCode</th><th class='input-valor'>Valor StarSoft </br> Item Pai</th><th>% Aumento</th><th class='input-valor'>Valor Complemento</th><th class='input-valor'>Valor M&iacute;nimo</th><th class='input-valor'>Valor StarSoft <br> Item M&iacute;nimo</th></tr>"			
				 response.write "<td><a href='https://brentwood.com.br/qrcode/index.php?search="  & rs("D41_001_C") & "' target='_blank'>" & rs("D41_001_C") & "</td>"
				 response.write "<td>" & rs("D04_008_C") & "</td>"
				
				querySqlServer = ""	
				querySqlServer = querySqlServer & " SELECT TOP 1 CASE WHEN D04.D04_506_N = 1 THEN CONVERT(DECIMAL(18,2), (D04.D04_003_B * D04.D04_011_B * 10 * ISNULL( ( ( (SELECT MAX(BW15.BW15_003_B) FROM BW15 WHERE BW15.UKEY IN (D04.BW15_UKEY, D04.BW15_UKEYA, D04.BW15_UKEYB, D04.BW15_UKEYC, D04.BW15_UKEYD, D04.BW15_UKEYE)) / 100) + 1), 1)) ) ELSE 0 END AS VLR FROM D04 (NOLOCK) WHERE D04.UKEY = ("
				querySqlServer = querySqlServer & " SELECT DBO.[FN_GETUKEYCATALOGO](d41.D04_UKEY) from d41 (NOLOCK) where D41_001_C = '" & trim(rs("D41_001_C")) & "' ) "
				
				 Set rsStar = connStar.Execute(querySqlServer)
				
				 If Not (rsStar.EOF And rsStar.BOF) Then
					 Do While Not rsStar.EOF
						 rsStar.MoveFirst
							 vlrStar = rsStar("VLR")									
						 rsStar.MoveNext			
					 Loop
				 end if
				
				TagPercent = ""
				
				if CDbl(vlrStar) > 0 and  CDbl(rs("VALOR")) > 0 then
					percent = FormatNumber ( ( ( CDbl(vlrStar) / CDbl(rs("VALOR")) ) - 1 ) * 100 ,2 )
					TagPercent = "<td><input class='input-percentual' name='pec[]' min='0' max='200000' step='0.0001' value='" & percent & " %' style='pointer-events:none;'></td>"
				else
					TagPercent = "<td><input class='input-percentual' name='pec[]' min='0' max='200000' step='0.0001' value='' style='pointer-events:none;'></td>"
				end if					
				
				vlrStar = Replace(Trim("" & vlrStar), ",", ".")
				vlrQrCode = Replace(Trim("" & rs("VALOR")), ",", ".")
				vlrComplemento = Replace(Trim("" & rs("VALORCOMPLEMENTO")), ",", ".")
				percent = 0
							
				 If CDbl(vlrStar) <> CDbl(vlrQrCode) Then
					 response.Write "<td><select onchange='habilitarCampo(this)' name='selecao[]' ><option value='0'>N&atilde;o</option><option value='1' selected>Sim</option></select></td>"
					 response.Write "<td><input class='input-valor' name='VALOR[]' min='0' max='200000' step='0.0001' value='" & vlrQrCode & "' disabled style='color:brown; background-color: yellow;'></td>"
					 response.Write "<td><input type='number' class='input-valor' name='vlrStar[]' min='0' max='200000' step='0.0001' value='" & vlrStar & "'  style='color:brown; background-color: yellow;'></td>"
					
					 response.Write TagPercent
					
					 response.Write "<td><input class='input-valor' name='vlrComplemento[]' min='0' max='200000' step='0.0001' value='" & vlrComplemento & "' style='color:brown; background-color: yellow;'></td>"
				 else
					 response.Write "<td><select onchange='habilitarCampo(this)' name='selecao[]'><option value='0' selected>N&atilde;o</option><option value='1'>Sim</option></select></td>"
					 response.write "<td><input class='input-valor' name='VALOR[]' min='0' max='200000' step='0.0001' value='" & vlrQrCode & "' disabled></td>"
					 response.write "<td><input type='number' class='input-valor' name='vlrStar[]' min='0' max='200000' step='0.0001' value='" & vlrStar & "' readonly></td>"
					
					 response.Write TagPercent
					
					 response.Write "<td><input class='input-valor' name='vlrComplemento[]' min='0' max='200000' step='0.0001' value='" & vlrComplemento & "' style='pointer-events:none;'></td>"
				 end if

				response.write "<td class='tdcenter'>" & Server.HTMLEncode("" & rs("VALORMINIMO")) & "</td>"

                ' Consulta o item referenciado por ITEMVALORMINIMO no StarSoft.
                vlrStarMinimo = 0
                referenciaMinimo = ""
                descricaoMinimo = ""
                If Not IsNull(rs("ITEMVALORMINIMO")) Then
                    referenciaMinimo = Trim(CStr(rs("ITEMVALORMINIMO")))
                End If

                If referenciaMinimo <> "" Then
                    queryStarMinimo = "SELECT TOP 1 CASE WHEN D04.D04_506_N = 1 THEN CONVERT(DECIMAL(18,2), (D04.D04_003_B * D04.D04_011_B * 10 * ISNULL(((SELECT MAX(BW15.BW15_003_B) FROM BW15 WHERE BW15.UKEY IN (D04.BW15_UKEY, D04.BW15_UKEYA, D04.BW15_UKEYB, D04.BW15_UKEYC, D04.BW15_UKEYD, D04.BW15_UKEYE)) / 100.0 + 1), 1))) ELSE 0 END AS VLR, D04.D04_008_C AS DESCRICAO_MINIMO FROM D04 (NOLOCK) WHERE D04.D04_001_C = '" & Replace(referenciaMinimo, "'", "''") & "'"

                    Set rsStarMinimo = connStar.Execute(queryStarMinimo)
                    If Not rsStarMinimo.EOF Then
                        If Not IsNull(rsStarMinimo("DESCRICAO_MINIMO")) Then
                            descricaoMinimo = CStr(rsStarMinimo("DESCRICAO_MINIMO"))
                        End If
                        If Not IsNull(rsStarMinimo("VLR")) Then
                            vlrStarMinimo = rsStarMinimo("VLR")
                        End If
                    End If
                    rsStarMinimo.Close
                    Set rsStarMinimo = Nothing
                End If

                vlrStarMinimo = Replace(Trim(CStr(vlrStarMinimo)), ",", ".")
                Response.Write "<td><input type='number' class='input-valor' name='vlrStarMinimo[]' min='0' max='200000' step='0.0001' value='" & vlrStarMinimo & "'></td>"

				response.write "<td><input type='hidden' name='etiqueta' value='" & rs("D41_001_C") & "'><input type='hidden' name='description[]' value='" & Server.HTMLEncode(descricaoMinimo) & "'></td>"
				response.write "</tr>"
				
				CODFORNECEDOR = rs("A08_001_C")
				vlrStar = 0
				vlrQrCode = 0
				
				rs.MoveNext
				
			end if	
			
		Do While Not rs.EOF
		
			vlrStar = 0
			contador = contador +1 
			
			if CODFORNECEDOR <> "" and CODFORNECEDOR <> rs("A08_001_C") then
				response.write "</table></fieldset><p align='center'><input type='submit' name='salvar' value='Efetuar Update' style='background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;'></p></form>"
				response.write "</br><form method='post' style='' action='update3.asp'>"
				response.write "<fieldset class='field_set'>"
				response.write "<legend><font size='2px' color='#9d46a3'>Rela&ccedil;&atilde;o de Produtos: " & rs("A08_001_C") & " - " & rs("A08_002_C") & "</font></legend>"	
				response.write "<table width='100%'><tr><th width='8%'>Etiqueta</th><th>Produto</th><th>Atualizar</th><th class='input-valor'>Valor QrCode</th><th class='input-valor'>Valor StarSoft </br> Item Pai</th><th>% Aumento</th><th class='input-valor'>Valor Complemento</th><th class='input-valor'>Valor M&iacute;nimo</th><th class='input-valor'>Valor StarSoft <br> Item M&iacute;nimo</th></tr>"	
			end if
		
			response.write "<td><a href='https://brentwood.com.br/qrcode/index.php?search="  & rs("D41_001_C" ) & "' target='_blank'>" & rs("D41_001_C") & "</td>"
			response.write "<td>" & rs("D04_008_C") & "-" & contador & "</td>"

					querySqlServer = ""
					querySqlServer = querySqlServer & " SELECT TOP 1 CASE WHEN D04.D04_506_N = 1 THEN CONVERT(DECIMAL(18,2), (D04.D04_003_B * D04.D04_011_B * 10 * ISNULL( ( ( (SELECT MAX(BW15.BW15_003_B) FROM BW15 WHERE BW15.UKEY IN (D04.BW15_UKEY, D04.BW15_UKEYA, D04.BW15_UKEYB, D04.BW15_UKEYC, D04.BW15_UKEYD, D04.BW15_UKEYE)) / 100) + 1), 1)) ) ELSE 0 END AS VLR FROM D04 (NOLOCK) WHERE D04.UKEY = ("
					querySqlServer = querySqlServer & " SELECT DBO.[FN_GETUKEYCATALOGO](d41.D04_UKEY) from d41 (NOLOCK) where D41_001_C = '" & trim(rs("D41_001_C")) & "' ) "
					
					Set rsStar = connStar.Execute(querySqlServer)						
					
					If Not (rsStar.EOF And rsStar.BOF) Then
						Do While Not rsStar.EOF
							 rsStar.MoveFirst
								vlrStar = rsStar("VLR")															 
							 rsStar.MoveNext			
						Loop
					end if				
					
					TagPercent = ""
					if CDbl(vlrStar) > 0 and CDbl(rs("VALOR")) > 0 then
						percent = FormatNumber ( ( ( CDbl(vlrStar) / CDbl(rs("VALOR")) ) - 1 ) * 100 ,2 )
						TagPercent = "<td><input class='input-percentual' name='pec[]' min='0' max='200000' step='0.0001' value='" & percent & " %' style='pointer-events:none;'></td>"
					else
						TagPercent = "<td><input class='input-percentual' name='pec[]' min='0' max='200000' step='0.0001' value='' style='pointer-events:none;'></td>"
					end if	
					
					vlrStar = Replace(Trim("" & vlrStar), ",", ".")
					vlrQrCode = Replace(Trim("" & rs("VALOR")), ",", ".")
					vlrComplemento = Replace(Trim("" & rs("VALORCOMPLEMENTO")), ",", ".")
										
					If CDbl(vlrStar) <> CDbl(vlrQrCode) Then
						response.Write "<td><select onchange='habilitarCampo(this)' name='selecao[]'><option value='0'>N&atilde;o</option><option value='1' selected>Sim</option></select></td>"
						response.Write "<td><input class='input-valor' name='VALOR[]' min='0' max='200000' step='0.0001' value='" & vlrQrCode & "' disabled style='color:brown; background-color: yellow;'></td>"
						response.Write "<td><input type='number' class='input-valor' name='vlrStar[]' min='0' max='200000' step='0.0001' value='" & vlrStar & "' style='color:brown; background-color: yellow;'></td>"
						
						response.Write TagPercent
						
						response.Write "<td><input class='input-valor' name='vlrComplemento[]' min='0' max='200000' step='0.0001' value='" & vlrComplemento & "' style='color:brown; background-color: yellow;'></td>"
					else
						response.Write "<td><select onchange='habilitarCampo(this)' name='selecao[]'><option value='0' selected>N&atilde;o</option><option value='1'>Sim</option></select></td>"
						response.write "<td><input class='input-valor' name='VALOR[]' min='0' max='200000' step='0.0001' value='" & vlrQrCode & "' disabled></td>"
						response.write "<td><input type='number' class='input-valor' name='vlrStar[]' min='0' max='200000' step='0.0001' value='" & vlrStar & "' readonly ></td>"
					
						response.Write TagPercent
						
						response.Write "<td><input class='input-valor' name='vlrComplemento[]' min='0' max='200000' step='0.0001' value='" & vlrComplemento & "' style='pointer-events:none;'></td>"
					end if
				
					response.write "<td class='tdcenter'>" & Server.HTMLEncode("" & rs("VALORMINIMO")) & "</td>"

                ' Consulta o item referenciado por ITEMVALORMINIMO no StarSoft.
                vlrStarMinimo = 0
                referenciaMinimo = ""
                descricaoMinimo = ""
                If Not IsNull(rs("ITEMVALORMINIMO")) Then
                    referenciaMinimo = Trim(CStr(rs("ITEMVALORMINIMO")))
                End If
				
                If referenciaMinimo <> "" Then
					queryStarMinimo = ""
					queryStarMinimo = queryStarMinimo & "SELECT TOP 1 "
					queryStarMinimo = queryStarMinimo & "CASE WHEN D04.D04_506_N = 1 THEN "
					queryStarMinimo = queryStarMinimo & "CONVERT(DECIMAL(18,2), ("
					queryStarMinimo = queryStarMinimo & "D04.D04_003_B * D04.D04_011_B * 10 * "
					queryStarMinimo = queryStarMinimo & "ISNULL(("
					queryStarMinimo = queryStarMinimo & "(SELECT MAX(BW15.BW15_003_B) "
					queryStarMinimo = queryStarMinimo & "FROM BW15 "
					queryStarMinimo = queryStarMinimo & "WHERE BW15.UKEY IN ("
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEY, "
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEYA, "
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEYB, "
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEYC, "
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEYD, "
					queryStarMinimo = queryStarMinimo & "D04.BW15_UKEYE"
					queryStarMinimo = queryStarMinimo & ")) / 100.0 + 1"
					queryStarMinimo = queryStarMinimo & "), 1)"
					queryStarMinimo = queryStarMinimo & ")) "
					queryStarMinimo = queryStarMinimo & "ELSE 0 END AS VLR, D04.D04_008_C AS DESCRICAO_MINIMO "
					queryStarMinimo = queryStarMinimo & "FROM D04 (NOLOCK) "
					queryStarMinimo = queryStarMinimo & "WHERE D04.D04_001_C = '"
					queryStarMinimo = queryStarMinimo & Replace(referenciaMinimo, "'", "''")
					queryStarMinimo = queryStarMinimo & "'"
					Set rsStarMinimo = connStar.Execute(queryStarMinimo)

					If Not rsStarMinimo.EOF Then
                        If Not IsNull(rsStarMinimo("DESCRICAO_MINIMO")) Then
                            descricaoMinimo = CStr(rsStarMinimo("DESCRICAO_MINIMO"))
                        End If
						If Not IsNull(rsStarMinimo("VLR")) Then
							vlrStarMinimo = rsStarMinimo("VLR")
						End If
					End If

					rsStarMinimo.Close
					Set rsStarMinimo = Nothing

				End If

                vlrStarMinimo = Replace(Trim(CStr(vlrStarMinimo)), ",", ".")
                Response.Write "<td><input type='number' class='input-valor' name='vlrStarMinimo[]' min='0' max='200000' step='0.0001' value='" & vlrStarMinimo & "'></td>"

					response.write "<td><input  type='hidden' name='etiqueta' value='" & rs("D41_001_C") & "'><input type='hidden' name='description[]' value='" & Server.HTMLEncode(descricaoMinimo) & "'></td>"
					response.write "</tr>"						
				
				CODFORNECEDOR = rs("A08_001_C")					
				querySqlServer = ""
				vlrStar = 0
				vlrQrCode = 0
				
			rs.MoveNext	
			
		Loop
				
			
		Else
			Response.Write "Nenhum registro encontrado: " &	 CODFORNECEDOR
		End If

		response.write "</table></fieldset><p align='center'><input type='submit' name='salvar' value='Efetuar Update' style='background: #7bd2e8; border-radius: 6px; padding: 10px; cursor: pointer; color: #fff; border: none; font-size: 16px;'></p></form>"
	
	end if

	%>			
			<div class="content">
			
			</div></br>
			
		</div>
	</div>	
	


</body></html>

	


</body></html>
