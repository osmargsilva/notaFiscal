# notaFiscal

function consultarNFe() {
  // Insira a chave de acesso da NF-e que você deseja consultar
  var chaveNFe = '12345678901234567890123456789012345678901234'; // Substitua pela chave real
  
  // URL para consulta da SEFAZ (consultar via consulta pública)
  var url = 'https://www.sefazvirtual.fazenda.gov.br/NFeConsulta/NFeConsulta4/NFeConsulta4.asmx?wsdl';
  
  // Montar o XML da consulta
  var xmlConsulta = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<consSitNFe xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">' +
      '<tpAmb>2</tpAmb>' +  // 1 para produção, 2 para homologação
      '<xServ>CONSULTAR</xServ>' +
      '<chNFe>' + chaveNFe + '</chNFe>' +
    '</consSitNFe>';

  // Configurar a requisição HTTP
  var options = {
    'method': 'post',
    'contentType': 'text/xml',
    'payload': xmlConsulta
  };
  
  // Enviar a requisição
  var resposta = UrlFetchApp.fetch(url, options);
  
  // Obter o conteúdo da resposta
  var xmlResponse = resposta.getContentText();
  
  // Parse do XML da resposta
  var document = XmlService.parse(xmlResponse);
  var root = document.getRootElement();
  
  // Extraindo dados específicos do XML da resposta
  var cStat = root.getChild('protNFe').getChild('infProt').getChild('cStat').getText();
  var descricao = root.getChild('protNFe').getChild('infProt').getChild('xMotivo').getText();
  
  // Se a consulta for bem-sucedida, vamos extrair mais detalhes
  if (cStat == '100') {  // Status "100" significa que a NF-e foi autorizada
    var numeroNota = root.getChild('protNFe').getChild('infProt').getChild('chNFe').getText();
    var dataEmissao = root.getChild('protNFe').getChild('infProt').getChild('dhRecbto').getText();
    var valorTotal = root.getChild('protNFe').getChild('infProt').getChild('vNF').getText();
    var cnpjEmitente = root.getChild('protNFe').getChild('infProt').getChild('CNPJ').getText();
    
    // Preencher as informações na planilha
    var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var linha = planilha.getLastRow() + 1;
    
    planilha.getRange(linha, 1).setValue(numeroNota);  // Número da nota
    planilha.getRange(linha, 2).setValue(dataEmissao);  // Data de emissão
    planilha.getRange(linha, 3).setValue(cnpjEmitente); // CNPJ do emitente
    planilha.getRange(linha, 4).setValue(valorTotal);   // Valor total da nota
  } else {
    Logger.log('Erro: ' + descricao);
  }
}
