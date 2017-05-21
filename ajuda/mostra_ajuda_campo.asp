<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../vbs/funcoes_interface.vbs" -->
<!--#include file="../vbs/dhm_variaveis.vbs" -->
<!--#include file="../inc/abre_conn.asp" -->
<%
'session("accesslevel") = 1
'  p_id_pagina,p_id_campo,p_obri,p_tx_obrig,p_caminho
Dim str_ID_PAGINA
Dim str_ID_CAMPO
Dim str_OBRIGATORIEDADE
Dim str_Tp_Ajuda
Dim strTitulo
strTitulo = "Tela de ajuda"
str_Tp_Ajuda = "C"
If Request.Form("p_id_pagina") <> "" then
	str_ID_PAGINA = Request.Form("p_id_pagina")
Else
	'Call sTrataErro(0,str_TxMensagem,"../")
End If
If Request.Form("p_id_campo") <> "" then
	str_ID_CAMPO = Request.Form("p_id_campo")
Else
	str_ID_CAMPO = 0
	str_Tp_Ajuda = "T"	
End If
If Request.Form("p_obri") <> "" then
	If Request.Form("p_obri") = 1 then
		str_OBRIGATORIEDADE = "Obrigatório"
	ElseIf Request.Form("p_obri") = 2 then
		str_OBRIGATORIEDADE = "não obrigatório"
	ElseIf Request.Form("p_obri") = 3 then
		str_OBRIGATORIEDADE = "Obrigatório dependente " & Request.Form("p_tx_obrig")
	End If
Else
	'str_TxMensagem = "Problema com o help deste campo"
	'Call sTrataErro(0,str_TxMensagem,"../")
End If
'1 - OBRIGATÓRIO 2 -  FACULTATIVO -  3 - OBRIGATÓRIO PENDENTE DE OUTRO
'set rds = Server.CreateObject("ADODB.Recordset")				
'set rds = conn_db.execute(strQuery)
'response.Write("aqui" & strQuery)
'response.End()
%>		
<html> 
<head> 
<TITLE><%=strTitulo%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/dhm.css" rel="stylesheet" type="text/css" media="screen">
<link href="../css/print.css" rel="stylesheet" type="text/css" media="print">
<script src="../js/funcoes.js" language="JavaScript"></script>
</head> 
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
'int_campo int_preencher
Dim rds_HELP_PAGINA_CAMPO
Dim vet_Desc_Help
Dim str_HELP_TX_TITULO_OBJETO
Dim str_HELP_TX_CONTEUDO
Dim str_HELP_TX_DESCRICAO
str_Sql = ""
str_Sql = str_Sql & " SELECT"
str_Sql = str_Sql & " HELP_PAGINA_CAMPO.HELP_TX_NOME_OBJ"
str_Sql = str_Sql & " , HELP_PAGINA_CAMPO.HELP_TX_TITULO_OBJETO"
str_Sql = str_Sql & " , HELP_PAGINA_CAMPO.HELP_TX_DESCRICAO"
str_Sql = str_Sql & " , HELP_PAGINA_CAMPO.HELP_TX_CONTEUDO"
str_Sql = str_Sql & " FROM HELP_PAGINA_CAMPO"
str_Sql = str_Sql & " WHERE"
str_Sql = str_Sql & " HELP_TX_IDENTIFICADOR_PAGINA = " & fTrataCampoSql(str_ID_PAGINA, "texto", "I","M") 
str_Sql = str_Sql & " AND HELP_TX_IDENTIFICADOR_CAMPO = " &  fTrataCampoSql(str_ID_CAMPO, "texto", "I","M")
'response.Write(str_Sql)
'response.End()
Set rds_HELP_PAGINA_CAMPO = Server.CreateObject("ADODB.Recordset")
rds_HELP_PAGINA_CAMPO.Open str_Sql, conn_db  
if not rds_HELP_PAGINA_CAMPO.Eof then 
	str_HELP_TX_TITULO_OBJETO = rds_HELP_PAGINA_CAMPO("HELP_TX_TITULO_OBJETO")							
	vet_Desc_Help = Split(rds_HELP_PAGINA_CAMPO("HELP_TX_DESCRICAO") , "/")				
	str_HELP_TX_CONTEUDO = rds_HELP_PAGINA_CAMPO("HELP_TX_CONTEUDO")							
	str_HELP_TX_DESCRICAO = ""
	For int_Indice = 0 to UBound(vet_Desc_Help)
		str_HELP_TX_DESCRICAO = str_HELP_TX_DESCRICAO & vet_Desc_Help(i) & "<p>"
		If vet_Desc_Help(i) = "" Then
			exit for
		End If
	Next
End If
rds_HELP_PAGINA_CAMPO.Close
set rds_HELP_PAGINA_CAMPO = Nothing
%>
<table width="532" height="300"  border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="100%" height="22"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td class="titulo_campo"> &gt;&gt; Ajuda</td>
        </tr>
      </table>
      <table width="100%"  border="0" cellspacing="4" cellpadding="0">
        <tr>
          <td width="11%">&nbsp;</td>
          <td width="89%">&nbsp;</td>
        </tr>
        <tr>
          <td class="lbl_item_12_e">Campo:</td>
          <td class="tabela_titulo_sem_ordem_c2"><%=str_HELP_TX_TITULO_OBJETO%></td>
        </tr>
      </table>
      <table width="100%"  border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td width="115">&nbsp;</td>
          <td width="341">&nbsp;</td>
        </tr>
        <tr>
          <td class="form_titulo_campo">Descrição:</td>
          <td class="form_campo_borda_esquerdo"><%=str_HELP_TX_DESCRICAO%></td>
        </tr>
        <tr>
          <td class="form_titulo_campo">Conteúdo:</td>
          <td class="form_campo_borda_esquerdo"><%=str_HELP_TX_CONTEUDO%></td>
        </tr>
        <tr>
          <td class="form_titulo_campo">Obrigatoriedade:</td>
          <td class="form_campo_borda_esquerdo"><%=str_OBRIGATORIEDADE%></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="100%"  border="0" cellspacing="4" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td class="arial_12_pt">&nbsp;</td>
        </tr>
        <tr>
          <td width="120"><div align="right" class="arial_12_pb"></div></td>
          <td width="410" class="arial_12_pt"><div align="right"><a href="JavaScript:window.close()"><img src="../botao_fechar.gif title="..:: Fechar janela" name="btnDesistir" width="85" height="19" border="0" id="btnDesistir" onClick="javascript:window.close();"></a></div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body> 
</html> 
<%
set conn_db = nothing
%>
