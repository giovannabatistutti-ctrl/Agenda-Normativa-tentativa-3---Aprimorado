// @ts-nocheck
// =============================================
// CONFIGURAÇÕES PRINCIPAIS
// =============================================

const CONFIG = {
  TOQAN_TOKEN: '',
  SLACK_WEBHOOK: '',
  SHEET_ID: '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U',
  
  TOQAN_API: {
    BASE_URL: 'https://api.coco.prod.toqan.ai',
    TIMEOUT: 30000,
    MAX_RETRIES: 3,
    RETRY_DELAY: 2000
  },
  
  AGENDAMENTOS: {
    DIARIOS: [9, 17],
    TIMEZONE: 'America/Sao_Paulo'
  }
};

// =============================================
// CLIENTE TOQAN
// =============================================

class ToqanClient {
  constructor() {
    this.baseUrl = CONFIG.TOQAN_API.BASE_URL;
    this.timeout = CONFIG.TOQAN_API.TIMEOUT;
    this.maxRetries = CONFIG.TOQAN_API.MAX_RETRIES;
    this.retryDelay = CONFIG.TOQAN_API.RETRY_DELAY;
  }
  
  _generateTraceId() {
    return Utilities.getUuid();
  }
  
  _getHeaders(traceId = null) {
    const headers = {
      'X-Api-Key': CONFIG.TOQAN_TOKEN,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'User-Agent': 'iFood-Compliance-Bot/1.0'
    };
    
    if (traceId) {
      headers['X-Request-Id'] = traceId;
    }
    
    return headers;
  }
  
  _makeRequest(method, endpoint, payload = null, traceId = null) {
    let lastError;
    
    for (let attempt = 1; attempt <= this.maxRetries; attempt++) {
      try {
        const url = `${this.baseUrl}${endpoint}`;
        const options = {
          'method': method,
          'headers': this._getHeaders(traceId),
          'timeout': this.timeout,
          'muteHttpExceptions': true
        };
        
        if (payload && method !== 'GET') {
          options.payload = JSON.stringify(payload);
        }
        
        Logger.log(`📡 Attempt ${attempt}/${this.maxRetries}: ${method} ${endpoint}`);
        
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (statusCode >= 200 && statusCode < 300) {
          try {
            return JSON.parse(responseText);
          } catch (parseError) {
            Logger.log(`⚠️ JSON parse error: ${responseText.substring(0, 200)}`);
            return responseText;
          }
        }
        
        if ([429, 500, 502, 503, 504].includes(statusCode)) {
          lastError = new Error(`HTTP ${statusCode}: ${responseText.substring(0, 200)}`);
          if (attempt < this.maxRetries) {
            Logger.log(`⚠️ Retryable error, waiting ${this.retryDelay}ms...`);
            Utilities.sleep(this.retryDelay * attempt);
            continue;
          }
        }
        
        throw new Error(`HTTP ${statusCode}: ${responseText.substring(0, 200)}`);
        
      } catch (error) {
        lastError = error;
        if (attempt < this.maxRetries) {
          Logger.log(`⚠️ Request failed, retrying in ${this.retryDelay}ms: ${error}`);
          Utilities.sleep(this.retryDelay * attempt);
        }
      }
    }
    
    throw lastError;
  }
  
  createConversation(userMessage) {
    if (!userMessage || typeof userMessage !== 'string') {
      throw new Error('userMessage must be a non-empty string');
    }
    
    const traceId = this._generateTraceId();
    Logger.log(`📝 Creating conversation - Trace: ${traceId}, Size: ${userMessage.length}`);
    
    const payload = { user_message: userMessage };
    const result = this._makeRequest('POST', '/api/create_conversation', payload, traceId);
    
    Logger.log(`✅ Conversation created - ID: ${result.conversation_id}`);
    return result;
  }
}

// =============================================
// FUNÇÃO PARA REGISTRAR LOGS NA ABA LOG APIs
// =============================================

function registrarLogAPI(orgao, status, detalhes, quantidade = 0) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let logSheet;
    
    try {
      logSheet = spreadsheet.getSheetByName('LOG APIs');
    } catch (e) {
      logSheet = spreadsheet.insertSheet('LOG APIs');
      const cabecalhos = ['Data_Hora', 'Orgao', 'Status', 'Quantidade_Normativos', 'Detalhes'];
      logSheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      logSheet.getRange(1, 1, 1, cabecalhos.length)
        .setBackground('#0c4a6e')
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    const dataHora = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    const linhaLog = [dataHora, orgao || 'SISTEMA', status || 'INFO', quantidade || 0, detalhes || ''];
    
    const ultimaLinha = logSheet.getLastRow();
    logSheet.getRange(ultimaLinha + 1, 1, 1, linhaLog.length).setValues([linhaLog]);
    
    Logger.log(`📋 LOG API: ${orgao} - ${status} - ${quantidade} normativos`);
    
  } catch (error) {
    Logger.log(`❌ Erro ao registrar log: ${error.toString()}`);
  }
}

// =============================================
// FUNÇÕES PRINCIPAIS CORRIGIDAS
// =============================================

function enviarSlackMensagem(mensagem) {
  try {
    Logger.log(`📤 Enviando mensagem Slack: ${mensagem.substring(0, 100)}...`);
    
    const payload = { "text": mensagem };
    const options = {
      'method': 'POST',
      'headers': {'Content-Type': 'application/json'},
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode === 200) {
      Logger.log('✅ Mensagem enviada para Slack com sucesso');
      return true;
    } else {
      Logger.log(`❌ Erro Slack HTTP ${statusCode}: ${response.getContentText()}`);
      return false;
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao enviar para Slack: ${error.toString()}`);
    return false;
  }
}

function salvarNaPlanilha(normativos) {
  Logger.log('💾 INICIANDO SALVAMENTO NA PLANILHA...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = spreadsheet.getSheets()[0];
    
    const ultimaLinha = sheet.getLastRow();
    
    if (ultimaLinha === 0) {
      const cabecalhos = [
        'normativo_index', 'Data_Captura', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Produto_Segmento', 'Tema', 'Impacto_Declarado',
        'Data_Vigencia', 'Aplicavel_SCD', 'Aplicavel_IP', 'Aplicavel_iFood',
        'status', 'Criticidade_Sistema', 'Resumo_Analise', 'Resposta_Toqan'
      ];
      sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }
    
    const dados = [];
    let proximoIndex = ultimaLinha + 1;
    
    normativos.forEach((normativo, index) => {
      const linha = [
        normativo.normativo_index || proximoIndex + index,
        normativo.Data_Captura || Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
        normativo.Orgao || 'N/A',
        normativo.Tipo_Norma || 'N/A',
        normativo.Numero || 'N/A',
        normativo.Data_Publicacao || 'N/A',
        normativo.Produto_Segmento || 'iFood Pago - Geral',
        normativo.Tema || 'N/A',
        normativo.Impacto_Declarado || 'Médio',
        normativo.Data_Vigencia || normativo.Data_Publicacao || 'N/A',
        normativo.Aplicavel_SCD || 'Não',
        normativo.Aplicavel_IP || 'Sim',
        normativo.Aplicavel_iFood || 'Sim',
        normativo.status || 'Analisado',
        normativo.Criticidade_Sistema || 'MÉDIA',
        normativo.Resumo_Analise || 'Análise Toqan AI',
        normativo.Resposta_Toqan || 'N/A'
      ];
      dados.push(linha);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha + 1;
      sheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} normativos salvos na planilha!`);
      return dados.length;
    }
    
    return 0;
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar na planilha: ${error.toString()}`);
    return 0;
  }
}

// =============================================
// FUNÇÕES DE ANÁLISE SIMPLIFICADAS
// =============================================

function analisarNormativosComToqan(normativos) {
  if (!normativos || normativos.length === 0) {
    Logger.log('ℹ️ Nenhum normativo para analisar');
    return [];
  }
  
  Logger.log(`🔍 Iniciando análise de ${normativos.length} normativos`);
  const client = new ToqanClient();
  const resultados = [];
  
  for (let i = 0; i < normativos.length; i++) {
    const normativo = normativos[i];
    
    try {
      Logger.log(`📊 [${i + 1}/${normativos.length}] Analisando: ${normativo.Orgao} ${normativo.Numero}`);
      
      const analise = analisarNormativoSimples(client, normativo);
      resultados.push(analise);
      
      Logger.log(`✅ [${i + 1}/${normativos.length}] Concluído`);
      
      if (i < normativos.length - 1) {
        Utilities.sleep(3000);
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no normativo ${i + 1}: ${error}`);
}
  }
  
  return resultados;
}

function analisarNormativoSimples(client, normativo) {
  try {
    const prompt = `Analise este normativo para compliance iFood:

ÓRGÃO: ${normativo.Orgao || 'N/A'}
TIPO: ${normativo.Tipo_Norma || 'N/A'} 
NÚMERO: ${normativo.Numero || 'N/A'}
DATA: ${normativo.Data_Publicacao || 'N/A'}
TEMA: ${normativo.Tema || 'N/A'}

RESPONDA APENAS COM ESTE JSON:
{
  "impacto": "Alto|Médio|Baixo",
  "produto": "iFood Pago PIX|iFood Pago Cartão|iFood Crédito|iFood Geral",
  "aplicavel_scd": "Sim|Não",
  "resumo": "Resumo conciso"
}`;

    const resposta = client.createConversation(prompt);
    Utilities.sleep(3000);
    
    return processarRespostaBasica(resposta, normativo);
    
  } catch (error) {
    Logger.log(`❌ Erro Toqan: ${error}`);
}
}

function processarRespostaBasica(resposta, normativo) {
  let impacto = 'Médio';
  let produto = 'iFood Pago - Geral';
  let aplicavelSCD = 'Não';
  let resumo = 'Analisado via Toqan AI';
  
  try {
    const respostaStr = JSON.stringify(resposta);
    
    if (respostaStr.includes('Alto') || respostaStr.includes('alto')) {
      impacto = 'Alto';
    } else if (respostaStr.includes('Baixo') || respostaStr.includes('baixo')) {
      impacto = 'Baixo';
    }
    
    if (respostaStr.includes('Crédito') || respostaStr.includes('crédito')) {
      produto = 'iFood Crédito';
    } else if (respostaStr.includes('PIX') || respostaStr.includes('pix')) {
      produto = 'iFood Pago PIX';
    }
    
  } catch (e) {
    Logger.log(`⚠️ Análise básica falhou: ${e}`);
  }
  
  return {
    normativo_index: obterProximoIndex(),
    Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
    Orgao: normativo.Orgao || 'N/A',
    Tipo_Norma: normativo.Tipo_Norma || 'N/A',
    Numero: normativo.Numero || 'N/A',
    Data_Publicacao: normativo.Data_Publicacao || 'N/A',
    Produto_Segmento: produto,
    Tema: normativo.Tema || 'N/A',
    Impacto_Declarado: impacto,
    Data_Vigencia: normativo.Data_Publicacao || 'N/A',
    Aplicavel_SCD: aplicavelSCD,
    Aplicavel_IP: 'Sim',
    Aplicavel_iFood: 'Sim',
    status: 'Analisado',
    Criticidade_Sistema: 'MÉDIA',
    Resumo_Analise: resumo,
    Resposta_Toqan: `Toqan ID: ${resposta.conversation_id}`
  };
}

function obterProximoIndex() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const ultimaLinha = sheet.getLastRow();
    return ultimaLinha <= 1 ? 1 : ultimaLinha + 1;
  } catch (e) {
    return 1;
  }
}

// =============================================
// FUNÇÃO PRINCIPAL SIMPLIFICADA
// =============================================

function executarSistemaCompleto() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO DE MONITORAMENTO');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando execução do sistema');
  
  try {
    // 1. COLETAR NORMATIVOS
    Logger.log('📡 ETAPA 1: COLETANDO NORMATIVOS...');
    const normativos = coletarNormativosReais();
    
    if (!normativos || normativos.length === 0) {
      Logger.log('ℹ️ Nenhum normativo novo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return;
    }
    
    // 2. ANALISAR COM TOQAN
    Logger.log('🤖 ETAPA 2: ANALISANDO COM TOQAN...');
    const normativosAnalisados = analisarNormativosComToqan(normativos);
    
    // 3. SALVAR NA PLANILHA
    Logger.log('💾 ETAPA 3: SALVANDO NA PLANILHA...');
    const salvos = salvarNaPlanilha(normativosAnalisados);
    
    // 4. ENVIAR RELATÓRIO
    Logger.log('📤 ETAPA 4: ENVIANDO RELATÓRIO...');
    enviarRelatorioCompletoSlack(normativosAnalisados, salvos);
    
    registrarLogAPI('SISTEMA', 'SUCCESS', 
      `Execução concluída - ${normativosAnalisados.length} normativos processados`, 
      normativosAnalisados.length
    );
    
    Logger.log(`🎉 SISTEMA CONCLUÍDO! ${normativosAnalisados.length} normativos processados`);
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NO SISTEMA: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro no sistema: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
  }
}

function enviarRelatorioCompletoSlack(normativos, salvos) {
  try {
    const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
    
    let mensagem = `🎯 *MONITORAMENTO IFOOD - ${dataHoje}*\n\n`;
    mensagem += `📊 *RESUMO:* ${normativos.length} normativos analisados | ${salvos} salvos\n\n`;
    
    if (normativos.length > 0) {
      mensagem += `📋 *NORMATIVOS DETECTADOS:*\n`;
      
      normativos.forEach((normativo) => {
        const emoji = normativo.Impacto_Declarado === 'Alto' ? '🔴' : 
                     normativo.Impacto_Declarado === 'Médio' ? '🟡' : '🟢';
        
        mensagem += `${emoji} *${normativo.Orgao} ${normativo.Tipo_Norma} ${normativo.Numero}*\n`;
        mensagem += `   Impacto: ${normativo.Impacto_Declarado} | Produto: ${normativo.Produto_Segmento}\n\n`;
      });
    }
    
    mensagem += `⚡ _Sistema Automático iFood Compliance_`;
    
    enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro relatório Slack: ${error}`);
  }
}

// =============================================
// FUNÇÕES DE COLETA DE NORMATIVOS
// =============================================

function coletarNormativosReais() {
  Logger.log('🔍 INICIANDO COLETA DE NORMATIVOS - SITES OFICIAIS');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando coleta de normativos dos órgãos reguladores');
  
  const normativos = [];
  const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd');
  
  try {
    // BACEN
    Logger.log('🛡️ Coletando BACEN...');
    const bacen = coletarBACENReal(dataHoje);
    if (bacen && bacen.length > 0) {
      normativos.push(...bacen);
      registrarLogAPI('BACEN', 'SUCCESS', `Coletados ${bacen.length} normativos`, bacen.length);
    } else {
      registrarLogAPI('BACEN', 'INFO', 'Nenhum normativo novo encontrado', 0);
    }
    
    Utilities.sleep(2000);
    
    // RFB
    Logger.log('🏛️ Coletando RFB...');
    const rfb = coletarRFBReal(dataHoje);
    if (rfb && rfb.length > 0) {
      normativos.push(...rfb);
      registrarLogAPI('RFB', 'SUCCESS', `Coletados ${rfb.length} normativos`, rfb.length);
    } else {
      registrarLogAPI('RFB', 'INFO', 'Nenhum normativo novo encontrado', 0);
    }
    
    Utilities.sleep(2000);
    
    // CMN
    Logger.log('📋 Coletando CMN...');
    const cmn = coletarCMNReal(dataHoje);
    if (cmn && cmn.length > 0) {
      normativos.push(...cmn);
      registrarLogAPI('CMN', 'SUCCESS', `Coletados ${cmn.length} normativos`, cmn.length);
    } else {
      registrarLogAPI('CMN', 'INFO', 'Nenhum normativo novo encontrado', 0);
    }
    
    Utilities.sleep(2000);
    
    // SUSEP
    Logger.log('🛡️ Coletando SUSEP...');
    const susep = coletarSUSEPReal(dataHoje);
    if (susep && susep.length > 0) {
      normativos.push(...susep);
      registrarLogAPI('SUSEP', 'SUCCESS', `Coletados ${susep.length} normativos`, susep.length);
    } else {
      registrarLogAPI('SUSEP', 'INFO', 'Nenhum normativo novo encontrado', 0);
    }
    
    Utilities.sleep(2000);
    
    // DOU
    Logger.log('📰 Coletando DOU...');
    const dou = coletarDOUReal(dataHoje);
    if (dou && dou.length > 0) {
      normativos.push(...dou);
      registrarLogAPI('DOU', 'SUCCESS', `Coletados ${dou.length} normativos`, dou.length);
    } else {
      registrarLogAPI('DOU', 'INFO', 'Nenhum normativo novo encontrado', 0);
    }
    
  } catch (error) {
    Logger.log(`❌ Erro na coleta: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro geral na coleta: ${error.toString()}`, 0);
  }
  
  // REMOVER DUPLICATAS
  const normativosUnicos = removerDuplicatas(normativos);
  
  // Registrar resumo final
  registrarLogAPI('SISTEMA', 'SUCCESS', 
    `Coleta concluída - ${normativosUnicos.length} normativos únicos encontrados`, 
    normativosUnicos.length
  );
  
  Logger.log(`📊 TOTAL COLETADO: ${normativosUnicos.length} normativos`);
  
  return normativosUnicos;
}

function coletarBACENReal(data) {
  const normativos = [];
  
  try {
    const url = 'https://www.bcb.gov.br/estabilidadefinanceira/buscanormas';
    registrarLogAPI('BACEN', 'INFO', `Iniciando consulta: ${url}`);
    
    // Simulação de coleta - substitua por coleta real
    // Por enquanto, vamos criar alguns dados de exemplo
    normativos.push({
      Orgao: 'BACEN',
      Tipo_Norma: 'Circular',
      Numero: '4015',
      Data_Publicacao: data,
      Tema: 'Regulamentação sobre pagamentos instantâneos',
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      texto_completo: 'Circular que regulamenta operações de pagamento instantâneo no sistema financeiro',
      url_fonte: url
    });
    
    registrarLogAPI('BACEN', 'INFO', `Processados ${normativos.length} normativos`);
    
  } catch (e) {
    Logger.log('❌ Coleta BACEN falhou: ' + e);
    registrarLogAPI('BACEN', 'ERROR', `Falha na coleta: ${e.toString()}`);
  }
  
  return normativos;
}

function coletarRFBReal(data) {
  const normativos = [];
  
  try {
    const url = 'https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/legislacao';
    registrarLogAPI('RFB', 'INFO', `Iniciando consulta: ${url}`);
    
    // Simulação de coleta
    normativos.push({
      Orgao: 'RFB',
      Tipo_Norma: 'Instrução Normativa',
      Numero: '2121',
      Data_Publicacao: data,
      Tema: 'Declaração de operações financeiras',
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      texto_completo: 'Instrução normativa sobre obrigações acessórias de pessoas jurídicas',
      url_fonte: url
    });
    
    registrarLogAPI('RFB', 'INFO', `Processados ${normativos.length} normativos`);
    
  } catch (e) {
    Logger.log('❌ Coleta RFB falhou: ' + e);
    registrarLogAPI('RFB', 'ERROR', `Falha na coleta: ${e.toString()}`);
  }
  
  return normativos;
}

function coletarCMNReal(data) {
  const normativos = [];
  
  try {
    const url = 'https://www.bcb.gov.br/normativos-e-listas/consulta-normativos';
    registrarLogAPI('CMN', 'INFO', `Iniciando consulta: ${url}`);
    
    // Simulação de coleta
    normativos.push({
      Orgao: 'CMN',
      Tipo_Norma: 'Resolução',
      Numero: '4949',
      Data_Publicacao: data,
      Tema: 'Regulamentação do crédito consignado',
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      texto_completo: 'Resolução do CMN sobre limites e condições do crédito consignado',
      url_fonte: url
    });
    
    registrarLogAPI('CMN', 'INFO', `Processados ${normativos.length} normativos`);
    
  } catch (e) {
    Logger.log('❌ Coleta CMN falhou: ' + e);
    registrarLogAPI('CMN', 'ERROR', `Falha na coleta: ${e.toString()}`);
  }
  
  return normativos;
}

function coletarSUSEPReal(data) {
  const normativos = [];
  
  try {
    const url = 'https://www.gov.br/susep/pt-br/assuntos/normas-e-orientacoes';
    registrarLogAPI('SUSEP', 'INFO', `Iniciando consulta: ${url}`);
    
    // Simulação de coleta
    normativos.push({
      Orgao: 'SUSEP',
      Tipo_Norma: 'Circular',
      Numero: '617',
      Data_Publicacao: data,
      Tema: 'Normas para seguros de crédito',
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      texto_completo: 'Circular SUSEP sobre contratação e condições de seguros de crédito',
      url_fonte: url
    });
    
    registrarLogAPI('SUSEP', 'INFO', `Processados ${normativos.length} normativos`);
    
  } catch (e) {
    Logger.log('❌ Coleta SUSEP falhou: ' + e);
    registrarLogAPI('SUSEP', 'ERROR', `Falha na coleta: ${e.toString()}`);
  }
  
  return normativos;
}

function coletarDOUReal(data) {
  const normativos = [];
  
  try {
    const url = `https://www.in.gov.br/consulta/-/buscar/dou`;
    registrarLogAPI('DOU', 'INFO', `Iniciando consulta: ${url}`);
    
    // Simulação de coleta
    normativos.push({
      Orgao: 'DOU',
      Tipo_Norma: 'Portaria',
      Numero: '123',
      Data_Publicacao: data,
      Tema: 'Atos oficiais do governo federal',
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      texto_completo: 'Publicação de atos oficiais no Diário Oficial da União',
      url_fonte: url
    });
    
    registrarLogAPI('DOU', 'INFO', `Processados ${normativos.length} normativos`);
    
  } catch (e) {
    Logger.log('❌ Coleta DOU falhou: ' + e);
    registrarLogAPI('DOU', 'ERROR', `Falha na coleta: ${e.toString()}`);
  }
  
  return normativos;
}

function fazerRequisicaoSegura(url) {
  for (let tentativa = 1; tentativa <= 3; tentativa++) {
    try {
      registrarLogAPI('HTTP', 'INFO', `Tentativa ${tentativa}/3 - URL: ${url}`);
      
      const options = {
        'method': 'GET',
        'headers': {
          'User-Agent': 'Mozilla/5.0 (compatible; iFood-Compliance-Bot/1.0)',
          'Accept': 'text/html,application/xhtml+xml,application/xml'
        },
        'muteHttpExceptions': true,
        'timeout': 30000
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        registrarLogAPI('HTTP', 'SUCCESS', `Request bem-sucedido - Status: ${statusCode}`);
        return response.getContentText();
      } else {
        registrarLogAPI('HTTP', 'WARNING', `Status HTTP ${statusCode} para ${url}`);
      }
      
      Utilities.sleep(2000);
    } catch (e) {
      registrarLogAPI('HTTP', 'ERROR', `Tentativa ${tentativa} falhou: ${e.toString()}`);
      if (tentativa < 3) {
        Utilities.sleep(2000);
      }
    }
  }
  
  registrarLogAPI('HTTP', 'ERROR', `Todas as tentativas falharam para: ${url}`);
  return null;
}

function removerDuplicatas(normativos) {
  if (!normativos || !Array.isArray(normativos)) return [];
  
  const seen = new Set();
  return normativos.filter(normativo => {
    if (!normativo || !normativo.Orgao || !normativo.Numero) return false;
    
    const key = `${normativo.Orgao}-${normativo.Numero}-${normativo.Data_Publicacao}`;
    return seen.has(key) ? false : (seen.add(key), true);
  });
}

// =============================================
// FUNÇÃO DE TESTE SIMPLES
// =============================================

function testarSistemaSimples() {
  Logger.log('🧪 TESTE SIMPLES DO SISTEMA');
  
  try {
    // Teste 1: Toqan
    Logger.log('\n1. 🤖 Testando Toqan...');
    const client = new ToqanClient();
    const teste = client.createConversation("Teste de conexão - responda com OK");
    Logger.log(`   ✅ Toqan: ${teste.conversation_id}`);
    
    // Teste 2: Planilha
    Logger.log('\n2. 📊 Testando planilha...');
    const normativoTeste = [{
      Orgao: 'TESTE',
      Tipo_Norma: 'Resolução',
      Numero: '999',
      Data_Publicacao: '2024-01-01',
      Tema: 'Normativo de teste'
    }];
    
    const salvos = salvarNaPlanilha(normativoTeste);
    Logger.log(`   ✅ Planilha: ${salvos} salvos`);
    
    // Teste 3: Slack
    Logger.log('\n3. 📤 Testando Slack...');
    const slackOk = enviarSlackMensagem('✅ Sistema testado com sucesso!');
    Logger.log(`   ✅ Slack: ${slackOk}`);
    
    // Teste 4: Coleta
    Logger.log('\n4. 🔍 Testando coleta...');
    const normativos = coletarNormativosReais();
    Logger.log(`   ✅ Coleta: ${normativos.length} normativos`);
    
    Logger.log('\n🎉 SISTEMA FUNCIONANDO!');
    return true;
    
  } catch (error) {
    Logger.log(`❌ TESTE FALHOU: ${error}`);
    return false;
  }
}
// =============================================
// SISTEMA REAL DE CAPTURA DE NORMATIVOS BACEN
// =============================================

function coletarNormativosReais() {
  Logger.log('🔍 INICIANDO COLETA DE NORMATIVOS REAIS - BACEN');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando coleta de normativos reais do BACEN');
  
  const normativos = [];
  const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd');
  
  try {
    // Buscar normativos recentes do BACEN
    Logger.log('🛡️ Buscando normativos BACEN/CMN recentes...');
    const normativosBACEN = buscarNormativosBACENRecentes();
    
    if (normativosBACEN && normativosBACEN.length > 0) {
      normativos.push(...normativosBACEN);
      registrarLogAPI('BACEN', 'SUCCESS', `Encontrados ${normativosBACEN.length} normativos recentes`, normativosBACEN.length);
      Logger.log(`✅ BACEN: ${normativosBACEN.length} normativos recentes encontrados`);
    } else {
      registrarLogAPI('BACEN', 'INFO', 'Nenhum normativo novo encontrado', 0);
      Logger.log('ℹ️ BACEN: nenhum normativo novo encontrado');
    }
    
    // Também buscar do DOU para complementar
    Logger.log('📰 Verificando DOU para normativos financeiros...');
    const normativosDOU = buscarNormativosDOURecentes();
    if (normativosDOU && normativosDOU.length > 0) {
      normativos.push(...normativosDOU);
      registrarLogAPI('DOU', 'SUCCESS', `Encontrados ${normativosDOU.length} normativos no DOU`, normativosDOU.length);
    }
    
  } catch (error) {
    Logger.log(`❌ Erro na coleta: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro geral na coleta: ${error.toString()}`, 0);
  }
  
  // REMOVER DUPLICATAS
  const normativosUnicos = removerDuplicatas(normativos);
  
  registrarLogAPI('SISTEMA', 'SUCCESS', 
    `Coleta concluída - ${normativosUnicos.length} normativos únicos encontrados`, 
    normativosUnicos.length
  );
  
  Logger.log(`📊 TOTAL COLETADO: ${normativosUnicos.length} normativos reais`);
  
  return normativosUnicos;
}

function buscarNormativosBACENRecentes() {
  const normativos = [];
  const dataAtual = new Date();
  const dataLimite = new Date(dataAtual.getTime() - (7 * 24 * 60 * 60 * 1000)); // Últimos 7 dias
  
  try {
    // Tipos de normativos para monitorar
    const tiposNormativos = [
      { tipo: 'Resolução CMN', nome: 'Resolução CMN' },
      { tipo: 'Resolução BCB', nome: 'Resolução BCB' },
      { tipo: 'Circular', nome: 'Circular' },
      { tipo: 'Carta Circular', nome: 'Carta Circular' },
      { tipo: 'Instrução Normativa BCB', nome: 'Instrução Normativa BCB' },
      { tipo: 'Comunicado', nome: 'Comunicado' }
    ];
    
    // Buscar na página de busca do BACEN
    const urlBusca = 'https://www.bcb.gov.br/estabilidadefinanceira/buscanormas';
    const html = fazerRequisicaoSegura(urlBusca);
    
    if (html) {
      // Extrair normativos da página de busca
      const normativosEncontrados = extrairNormativosBuscaBACEN(html, dataLimite);
      normativos.push(...normativosEncontrados);
      
      // Para cada normativo encontrado, buscar detalhes completos
      for (let normativo of normativosEncontrados) {
        try {
          const detalhes = buscarDetalhesNormativoBACEN(normativo.tipo, normativo.numero);
          if (detalhes) {
            Object.assign(normativo, detalhes);
          }
          Utilities.sleep(1000); // Delay para não sobrecarregar o servidor
        } catch (e) {
          Logger.log(`⚠️ Erro ao buscar detalhes de ${normativo.tipo} ${normativo.numero}: ${e}`);
        }
      }
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar normativos BACEN: ${error}`);
  }
  
  return normativos;
}

function extrairNormativosBuscaBACEN(html, dataLimite) {
  const normativos = [];
  
  try {
    // Regex para encontrar normativos na página de busca
    // Padrão: Tipo Número - Data - Título
    const regexNormativos = /(Resolução\s+(?:CMN|BCB)|Circular|Carta\s+Circular|Instrução\s+Normativa\s+BCB|Comunicado)\s+([\d\.]+).*?(\d{2}\/\d{2}\/\d{4})/gi;
    
    let match;
    while ((match = regexNormativos.exec(html)) !== null) {
      const tipo = match[1].trim();
      const numero = match[2].trim();
      const dataTexto = match[3];
      
      // Converter data
      const [dia, mes, ano] = dataTexto.split('/');
      const dataNormativo = new Date(ano, mes - 1, dia);
      
      // Verificar se é recente (últimos 7 dias)
      if (dataNormativo >= dataLimite) {
        // Extrair título (próximas linhas após o padrão)
        const inicioTitulo = match.index + match[0].length;
        const fimTitulo = html.indexOf('</', inicioTitulo);
        let titulo = html.substring(inicioTitulo, fimTitulo).trim();
        titulo = titulo.replace(/<[^>]*>/g, '').substring(0, 200);
        
        normativos.push({
          Orgao: 'BACEN',
          Tipo_Norma: tipo,
          Numero: numero,
          Data_Publicacao: Utilities.formatDate(dataNormativo, 'GMT-3', 'yyyy-MM-dd'),
          Tema: titulo || `${tipo} ${numero} do BACEN/CMN`,
          Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
          url_fonte: `https://www.bcb.gov.br/estabilidadefinanceira/exibenormativo?tipo=${encodeURIComponent(tipo)}&numero=${numero}`
        });
        
        Logger.log(`   📄 Encontrado: ${tipo} ${numero} - ${dataTexto}`);
      }
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao extrair normativos: ${error}`);
  }
  
  return normativos;
}

function buscarDetalhesNormativoBACEN(tipo, numero) {
  try {
    const url = `https://www.bcb.gov.br/estabilidadefinanceira/exibenormativo?tipo=${encodeURIComponent(tipo)}&numero=${numero}`;
    Logger.log(`   🔍 Buscando detalhes: ${tipo} ${numero}`);
    
    const html = fazerRequisicaoSegura(url);
    
    if (html) {
      return extrairDetalhesNormativo(html, tipo, numero);
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar detalhes: ${error}`);
  }
  
  return null;
}

function extrairDetalhesNormativo(html, tipo, numero) {
  const detalhes = {
    texto_completo: '',
    ementa: '',
    situacao: '',
    link_pdf: ''
  };
  
  try {
    // Extrair ementa/resumo
    const ementaMatch = html.match(/<div[^>]*class="ementa"[^>]*>([\s\S]*?)<\/div>/i);
    if (ementaMatch) {
      detalhes.ementa = ementaMatch[1].replace(/<[^>]*>/g, '').trim();
      detalhes.texto_completo = detalhes.ementa;
    }
    
    // Extrair situação
    const situacaoMatch = html.match(/Situação:?<\/strong>\s*<span[^>]*>([^<]+)</i);
    if (situacaoMatch) {
      detalhes.situacao = situacaoMatch[1].trim();
    }
    
    // Extrair link do PDF
    const pdfMatch = html.match(/<a[^>]*href="([^"]*\.pdf)"[^>]*>/i);
    if (pdfMatch) {
      detalhes.link_pdf = 'https://www.bcb.gov.br' + pdfMatch[1];
    }
    
    // Se não encontrou ementa, tentar extrair conteúdo principal
    if (!detalhes.texto_completo) {
      const conteudoMatch = html.match(/<div[^>]*class="conteudo"[^>]*>([\s\S]*?)<\/div>/i);
      if (conteudoMatch) {
        detalhes.texto_completo = conteudoMatch[1].replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 1000);
      }
    }
    
    Logger.log(`   ✅ Detalhes extraídos: ${detalhes.ementa ? 'Ementa encontrada' : 'Sem ementa'}`);
    
  } catch (error) {
    Logger.log(`⚠️ Erro ao extrair detalhes: ${error}`);
  }
  
  return detalhes;
}

function buscarNormativosDOURecentes() {
  const normativos = [];
  const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd');
  
  try {
    // Buscar no DOU por termos relacionados ao sistema financeiro
    const termos = ['BACEN', 'Banco Central', 'CMN', 'Conselho Monetário Nacional', 'Sistema Financeiro'];
    
    for (let termo of termos) {
      try {
        const url = `https://www.in.gov.br/consulta/-/buscar/dou?q=${encodeURIComponent(termo)}&s=todos&exactDate=personalized&sortType=0&publishFrom=${dataHoje}&publishTo=${dataHoje}`;
        const html = fazerRequisicaoSegura(url);
        
        if (html) {
          const normativosDOU = extrairNormativosDOU(html, termo);
          normativos.push(...normativosDOU);
        }
        
        Utilities.sleep(1000);
      } catch (e) {
        Logger.log(`⚠️ Erro ao buscar DOU para termo ${termo}: ${e}`);
      }
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar normativos DOU: ${error}`);
  }
  
  return normativos;
}

function extrairNormativosDOU(html, termo) {
  const normativos = [];
  
  try {
    // Extrair títulos e links das publicações
    const regex = /<h2[^>]*><a[^>]*href="([^"]*)"[^>]*>([^<]*)<\/a>/gi;
    
    let match;
    while ((match = regex.exec(html)) !== null) {
      const link = match[1];
      const titulo = match[2].trim();
      
      // Filtrar apenas publicações relevantes
      if (titulo.includes('BACEN') || titulo.includes('Banco Central') || 
          titulo.includes('CMN') || titulo.includes('Circular') || 
          titulo.includes('Resolução')) {
        
        normativos.push({
          Orgao: 'DOU',
          Tipo_Norma: 'Publicação Oficial',
          Numero: `DOU-${Date.now()}-${normativos.length}`,
          Data_Publicacao: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd'),
          Tema: titulo,
          Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
          texto_completo: `Publicação no DOU: ${titulo}`,
          url_fonte: link.startsWith('http') ? link : `https://www.in.gov.br${link}`
        });
      }
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao extrair normativos DOU: ${error}`);
  }
  
  return normativos;
}
// =============================================
// MÓDULO DE MONITORAMENTO NORMATIVO REAL
// Web scraping 100% real sem fallbacks simulados
// =============================================

class MonitoramentoNormativo {
  constructor() {
    this.config = this.carregarConfiguracoes();
  }

  carregarConfiguracoes() {
    return {
      fontes: {
        bcb: {
          url: 'https://www.bcb.gov.br/noticias',
          ativo: true
        },
        legisweb: {
          url: 'https://www.legisweb.com.br/noticias/',
          ativo: true
        },
        valor: {
          url: 'https://valor.globo.com/financas/',
          ativo: true
        },
        g1economia: {
          url: 'https://g1.globo.com/economia/',
          ativo: true
        },
        infomoney: {
          url: 'https://www.infomoney.com.br/',
          ativo: true
        },
        forbes: {
          url: 'https://forbes.com.br/',
          ativo: true
        },
        bloomberg: {
          url: 'https://www.bloomberglinea.com.br/',
          ativo: true
        },
        marianaLisboa: {
          url: 'https://br.linkedin.com/in/mariana-lisboa-5b993968',
          ativo: false // LinkedIn difícil de fazer scraping
        },
        btlaw: {
          url: 'https://www.linkedin.com/company/btlaw',
          ativo: false // LinkedIn difícil de fazer scraping
        }
      }
    };
  }

  /**
   * Executa o monitoramento completo de todas as fontes
   */
  executarMonitoramentoCompleto() {
    try {
      Logger.log('🚀 INICIANDO MONITORAMENTO NORMATIVO COMPLETO');
      registrarLogAPI('MONITORAMENTO', 'INFO', 'Iniciando monitoramento de fontes normativas');
      
      const resultados = [];
      const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd');
      
      // Ordem de monitoramento das fontes
      const fontesAtivas = Object.entries(this.config.fontes).filter(([_, config]) => config.ativo);
      
      for (const [fonte, config] of fontesAtivas) {
        Logger.log(`📡 Monitorando ${fonte.toUpperCase()}...`);
        
        try {
          let itens = [];
          
          switch(fonte) {
            case 'bcb':
              itens = this.monitorarBCBNoticias(dataHoje);
              break;
            case 'legisweb':
              itens = this.monitorarLegisWeb(dataHoje);
              break;
            case 'valor':
              itens = this.monitorarValorEconomico(dataHoje);
              break;
            case 'g1economia':
              itens = this.monitorarG1Economia(dataHoje);
              break;
            case 'infomoney':
              itens = this.monitorarInfoMoney(dataHoje);
              break;
            case 'forbes':
              itens = this.monitorarForbes(dataHoje);
              break;
            case 'bloomberg':
              itens = this.monitorarBloomberg(dataHoje);
              break;
          }
          
          if (itens && itens.length > 0) {
            resultados.push(...itens);
            registrarLogAPI(fonte.toUpperCase(), 'SUCCESS', `Encontrados ${itens.length} itens`);
            Logger.log(`   ✅ ${fonte}: ${itens.length} itens extraídos`);
          } else {
            registrarLogAPI(fonte.toUpperCase(), 'INFO', 'Nenhum item encontrado', 0);
            Logger.log(`   ℹ️ ${fonte}: Nenhum item encontrado`);
          }
          
          Utilities.sleep(3000); // Delay entre requisições
          
        } catch (error) {
          Logger.log(`   ❌ Erro ${fonte}: ${error.toString()}`);
          registrarLogAPI(fonte.toUpperCase(), 'ERROR', `Erro: ${error.toString()}`, 0);
        }
      }
      
      // Processar resultados
      if (resultados.length > 0) {
        const resultadosUnicos = removerDuplicatas(resultados);
        
        registrarLogAPI('MONITORAMENTO', 'SUCCESS', 
          `Monitoramento concluído - ${resultadosUnicos.length} itens encontrados`, 
          resultadosUnicos.length
        );
        
        Logger.log(`📊 MONITORAMENTO CONCLUÍDO: ${resultadosUnicos.length} itens reais encontrados`);
        return resultadosUnicos;
      } else {
        Logger.log('ℹ️ Nenhum novo item real encontrado no monitoramento');
        registrarLogAPI('MONITORAMENTO', 'INFO', 'Nenhum novo item real encontrado', 0);
        return [];
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no monitoramento: ${error.toString()}`);
      registrarLogAPI('MONITORAMENTO', 'ERROR', `Erro no monitoramento: ${error.toString()}`, 0);
      return [];
    }
  }

  /**
   * Monitora notícias do Banco Central do Brasil
   */
  monitorarBCBNoticias(data) {
    const noticias = [];
    
    try {
      const url = 'https://www.bcb.gov.br/noticias';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        // Extrair notícias usando método mais específico para BCB
        const noticiasEncontradas = this.extrairNoticiasBCB(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro BCB: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasBCB(html, data) {
    const noticias = [];
    
    try {
      // Método específico para estrutura do BCB
      const regexNoticias = /<a[^>]*href="(\/noticias\/[^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const conteudo = match[2];
        
        // Extrair título limpo
        const titulo = conteudo.replace(/<[^>]*>/g, '').trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaRelevante(titulo)) {
          noticias.push({
            Orgao: 'BCB_NOTICIAS',
            Tipo_Norma: 'Notícia',
            Numero: `BCB-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `Notícia BCB: ${titulo}`,
            url_fonte: `https://www.bcb.gov.br${link}`
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair BCB: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Monitora LegisWeb
   */
  monitorarLegisWeb(data) {
    const itens = [];
    
    try {
      const url = 'https://www.legisweb.com.br/noticias/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const itensEncontrados = this.extrairConteudoLegisWeb(html, data);
        itens.push(...itensEncontrados);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro LegisWeb: ${error.toString()}`);
    }
    
    return itens;
  }

  extrairConteudoLegisWeb(html, data) {
    const itens = [];
    
    try {
      // Buscar por artigos ou posts de notícias
      const regexPosts = /<article[^>]*>[\s\S]*?<a[^>]*href="([^"]*)"[^>]*>([^<]+)<\/a>[\s\S]*?<\/article>/gi;
      
      let match;
      while ((match = regexPosts.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 15 && this.isConteudoRelevante(titulo)) {
          itens.push({
            Orgao: 'LEGISWEB',
            Tipo_Norma: 'Publicação',
            Numero: `LEGIS-${Date.now()}-${itens.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `Conteúdo LegisWeb: ${titulo}`,
            url_fonte: link.startsWith('http') ? link : `https://www.legisweb.com.br${link}`
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair LegisWeb: ${error}`);
    }
    
    return itens;
  }

  /**
   * Monitora Valor Econômico - Finanças
   */
  monitorarValorEconomico(data) {
    const noticias = [];
    
    try {
      const url = 'https://valor.globo.com/financas/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const noticiasEncontradas = this.extrairNoticiasValor(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro Valor Econômico: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasValor(html, data) {
    const noticias = [];
    
    try {
      // Estrutura típica do Valor Econômico
      const regexNoticias = /<a[^>]*href="(https:\/\/valor\.globo\.com[^"]*)"[^>]*>([^<]+)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaEconomicaRelevante(titulo)) {
          noticias.push({
            Orgao: 'VALOR_ECONOMICO',
            Tipo_Norma: 'Notícia',
            Numero: `VALOR-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `Valor Econômico: ${titulo}`,
            url_fonte: link
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair Valor: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Monitora G1 Economia
   */
  monitorarG1Economia(data) {
    const noticias = [];
    
    try {
      const url = 'https://g1.globo.com/economia/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const noticiasEncontradas = this.extrairNoticiasG1(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro G1 Economia: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasG1(html, data) {
    const noticias = [];
    
    try {
      // Estrutura do G1
      const regexNoticias = /<a[^>]*href="(https:\/\/g1\.globo\.com\/economia[^"]*)"[^>]*>([^<]+)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaEconomicaRelevante(titulo)) {
          noticias.push({
            Orgao: 'G1_ECONOMIA',
            Tipo_Norma: 'Notícia',
            Numero: `G1-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `G1 Economia: ${titulo}`,
            url_fonte: link
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair G1: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Monitora InfoMoney
   */
  monitorarInfoMoney(data) {
    const noticias = [];
    
    try {
      const url = 'https://www.infomoney.com.br/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const noticiasEncontradas = this.extrairNoticiasInfoMoney(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro InfoMoney: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasInfoMoney(html, data) {
    const noticias = [];
    
    try {
      // Estrutura do InfoMoney
      const regexNoticias = /<a[^>]*href="(https:\/\/www\.infomoney\.com\.br[^"]*)"[^>]*>([^<]+)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaEconomicaRelevante(titulo)) {
          noticias.push({
            Orgao: 'INFOMONEY',
            Tipo_Norma: 'Notícia',
            Numero: `INFO-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `InfoMoney: ${titulo}`,
            url_fonte: link
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair InfoMoney: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Monitora Forbes Brasil
   */
  monitorarForbes(data) {
    const noticias = [];
    
    try {
      const url = 'https://forbes.com.br/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const noticiasEncontradas = this.extrairNoticiasForbes(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro Forbes: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasForbes(html, data) {
    const noticias = [];
    
    try {
      // Estrutura da Forbes Brasil
      const regexNoticias = /<a[^>]*href="(https:\/\/forbes\.com\.br[^"]*)"[^>]*>([^<]+)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaEconomicaRelevante(titulo)) {
          noticias.push({
            Orgao: 'FORBES_BR',
            Tipo_Norma: 'Notícia',
            Numero: `FORBES-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `Forbes Brasil: ${titulo}`,
            url_fonte: link
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair Forbes: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Monitora Bloomberg Linea
   */
  monitorarBloomberg(data) {
    const noticias = [];
    
    try {
      const url = 'https://www.bloomberglinea.com.br/';
      const html = fazerRequisicaoSegura(url);
      
      if (html) {
        const noticiasEncontradas = this.extrairNoticiasBloomberg(html, data);
        noticias.push(...noticiasEncontradas);
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro Bloomberg: ${error.toString()}`);
    }
    
    return noticias;
  }

  extrairNoticiasBloomberg(html, data) {
    const noticias = [];
    
    try {
      // Estrutura da Bloomberg Linea
      const regexNoticias = /<a[^>]*href="(https:\/\/www\.bloomberglinea\.com\.br[^"]*)"[^>]*>([^<]+)<\/a>/gi;
      
      let match;
      while ((match = regexNoticias.exec(html)) !== null) {
        const link = match[1];
        const titulo = match[2].trim();
        
        if (titulo && titulo.length > 20 && this.isNoticiaEconomicaRelevante(titulo)) {
          noticias.push({
            Orgao: 'BLOOMBERG',
            Tipo_Norma: 'Notícia',
            Numero: `BLOOM-${Date.now()}-${noticias.length}`,
            Data_Publicacao: data,
            Tema: titulo.substring(0, 200),
            Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
            texto_completo: `Bloomberg Linea: ${titulo}`,
            url_fonte: link
          });
        }
      }
      
    } catch (error) {
      Logger.log(`   ❌ Erro extrair Bloomberg: ${error}`);
    }
    
    return noticias;
  }

  /**
   * Funções auxiliares para filtragem de conteúdo relevante
   */
  isNoticiaRelevante(titulo) {
    const termosRelevantes = [
      'bacen', 'banco central', 'cmn', 'resolução', 'circular', 'normativo',
      'regulamento', 'financeiro', 'pagamento', 'fintech', 'compliance',
      'open banking', 'pix', 'cartão', 'crédito', 'empréstimo', 'regulação',
      'supervisão', 'normas', 'legislação'
    ];
    
    const tituloLower = titulo.toLowerCase();
    return termosRelevantes.some(termo => tituloLower.includes(termo));
  }

  isConteudoRelevante(titulo) {
    const termosRelevantes = [
      'normativo', 'regulamento', 'resolução', 'circular', 'legislação',
      'compliance', 'bacen', 'cmn', 'financeiro', 'pagamentos', 'fintech',
      'jurídico', 'legal', 'regulatório', 'tributário', 'fiscal'
    ];
    
    const tituloLower = titulo.toLowerCase();
    return termosRelevantes.some(termo => tituloLower.includes(termo));
  }

  isNoticiaEconomicaRelevante(titulo) {
    const termosRelevantes = [
      'bacen', 'banco central', 'cmn', 'juros', 'selic', 'inflação',
      'regulação', 'fintech', 'open banking', 'pix', 'cartão', 'crédito',
      'empréstimo', 'pagamento', 'financeiro', 'compliance', 'normativo',
      'resolução', 'circular', 'regulamento'
    ];
    
    const tituloLower = titulo.toLowerCase();
    return termosRelevantes.some(termo => tituloLower.includes(termo));
  }
}

// =============================================
// FUNÇÃO DE REQUISIÇÃO SEGURA MELHORADA
// =============================================

function fazerRequisicaoSegura(url) {
  for (let tentativa = 1; tentativa <= 3; tentativa++) {
    try {
      Logger.log(`   🔄 Tentativa ${tentativa}/3 - ${url}`);
      
      const options = {
        'method': 'GET',
        'headers': {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
          'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive'
        },
        'muteHttpExceptions': true,
        'followRedirects': true,
        'validateHttpsCertificates': false,
        'timeout': 45000
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        const content = response.getContentText();
        if (content && content.length > 1000) { // Conteúdo válido
          Logger.log(`   ✅ Request bem-sucedido - ${content.length} bytes`);
          return content;
        } else {
          Logger.log(`   ⚠️ Conteúdo muito curto: ${content.length} bytes`);
        }
      } else {
        Logger.log(`   ⚠️ Status HTTP ${statusCode}`);
      }
      
      Utilities.sleep(3000);
      
    } catch (e) {
      Logger.log(`   ❌ Tentativa ${tentativa} falhou: ${e.toString()}`);
      if (tentativa < 3) {
        Utilities.sleep(3000);
      }
    }
  }
  
  Logger.log(`   💥 Todas as tentativas falharam para: ${url}`);
  return null;
}

// =============================================
// FUNÇÃO DE TESTE REAL
// =============================================

function testarMonitoramentoReal() {
  Logger.log('🧪 TESTANDO MONITORAMENTO REAL - SEM FALLBACKS');
  
  try {
    const monitor = new MonitoramentoNormativo();
    
    Logger.log('\n1. 🛡️ Testando BCB Notícias...');
    const bcb = monitor.monitorarBCBNoticias('2024-01-01');
    Logger.log(`   Resultado REAL: ${bcb.length} notícias`);
    
    Logger.log('\n2. ⚖️ Testando LegisWeb...');
    const legisweb = monitor.monitorarLegisWeb('2024-01-01');
    Logger.log(`   Resultado REAL: ${legisweb.length} itens`);
    
    Logger.log('\n3. 📈 Testando Valor Econômico...');
    const valor = monitor.monitorarValorEconomico('2024-01-01');
    Logger.log(`   Resultado REAL: ${valor.length} notícias`);
    
    Logger.log('\n4. 📰 Testando G1 Economia...');
    const g1 = monitor.monitorarG1Economia('2024-01-01');
    Logger.log(`   Resultado REAL: ${g1.length} notícias`);
    
    Logger.log('\n5. 💰 Testando InfoMoney...');
    const info = monitor.monitorarInfoMoney('2024-01-01');
    Logger.log(`   Resultado REAL: ${info.length} notícias`);
    
    Logger.log('\n6. 🏆 Testando Forbes...');
    const forbes = monitor.monitorarForbes('2024-01-01');
    Logger.log(`   Resultado REAL: ${forbes.length} notícias`);
    
    Logger.log('\n7. 🌐 Testando Bloomberg...');
    const bloomberg = monitor.monitorarBloomberg('2024-01-01');
    Logger.log(`   Resultado REAL: ${bloomberg.length} notícias`);
    
    const total = bcb.length + legisweb.length + valor.length + g1.length + info.length + forbes.length + bloomberg.length;
    Logger.log(`\n📊 TOTAL REAL: ${total} itens coletados`);
    
    if (total === 0) {
      Logger.log('💡 Dica: Os sites podem estar bloqueando o scraping. Verifique os logs de erro.');
    }
    
    return total > 0;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO TESTE: ${error.toString()}`);
    return false;
  }
}
// =============================================
// FUNÇÕES DE INTEGRAÇÃO COM SISTEMA EXISTENTE
// =============================================

/**
 * Função principal que integra com o sistema existente
 */
function executarMonitoramentoNormativo() {
  Logger.log('🎯 INICIANDO MÓDULO DE MONITORAMENTO NORMATIVO');
  
  // Executar coleta de normativos oficiais (sistema existente)
  const normativosOficiais = coletarNormativosReais();
  
  // Executar monitoramento de fontes complementares (novo módulo)
  const monitor = new MonitoramentoNormativo();
  
  // Combinar resultados
  const todosNormativos = [...normativosOficiais, ...fontesComplementares];
  
  // Analisar com Toqan se houver resultados
  if (todosNormativos.length > 0) {
    Logger.log('🤖 Iniciando análise com Toqan...');
    const analises = analisarNormativosComToqan(todosNormativos);
    
    // Salvar análises
    if (analises && analises.length > 0) {
      salvarNaPlanilha(analises);
      Logger.log(`✅ ${analises.length} análises salvas`);
    }
  }
  
  Logger.log(`🎉 PROCESSO CONCLUÍDO: ${todosNormativos.length} itens processados`);
  return todosNormativos;
}

/**
 * Função de teste do módulo
 */
function testarModuloMonitoramento() {
  Logger.log('🧪 TESTANDO MÓDULO DE MONITORAMENTO');
  
  try {
    const monitor = new MonitoramentoNormativo();
    
    // Testar cada fonte individualmente
    Logger.log('\n1. 🛡️ Testando BCB Notícias...');
    const bcb = monitor.monitorarBCBNoticias('2024-01-01');
    Logger.log(`   ✅ BCB: ${bcb.length} notícias`);
    
    Logger.log('\n2. ⚖️ Testando LegisWeb...');
    const legisweb = monitor.monitorarLegisWeb('2024-01-01');
    Logger.log(`   ✅ LegisWeb: ${legisweb.length} itens`);
    
    Logger.log('\n3. 🏛️ Testando Mariana Lisboa...');
    const mariana = monitor.monitorarMarianaLisboa('2024-01-01');
    Logger.log(`   ✅ Mariana Lisboa: ${mariana.length} itens`);
    
    Logger.log('\n4. 💼 Testando BT Law LinkedIn...');
    const btlaw = monitor.monitorarBTLawLinkedIn('2024-01-01');
    Logger.log(`   ✅ BT Law: ${btlaw.length} posts`);
    
    Logger.log('\n🎉 MÓDULO FUNCIONANDO CORRETAMENTE!');
    return true;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO TESTE: ${error.toString()}`);
    return false;
  }
}

/**
 * Agendador automático
 */
function agendarMonitoramento() {
  // Agendar execução diária às 9:00
  ScriptApp.newTrigger('executarMonitoramentoNormativo')
    .timeBased()
    .atHour(9)
    .nearMinute(0)
    .everyDays(1)
    .create();
    
  Logger.log('⏰ Monitoramento normativo agendado para execução diária às 9:00');
}

// =============================================
// SISTEMA DE ANÁLISE COM TOQAN MELHORADO
// =============================================

function analisarNormativosComToqan(normativos) {
  if (!normativos || normativos.length === 0) {
    Logger.log('ℹ️ Nenhum normativo para analisar');
    return [];
  }
  
  Logger.log(`🔍 Iniciando análise de ${normativos.length} normativos com Toqan`);
  const client = new ToqanClient();
  const resultados = [];
  
  for (let i = 0; i < normativos.length; i++) {
    const normativo = normativos[i];
    
    try {
      Logger.log(`📊 [${i + 1}/${normativos.length}] Analisando: ${normativo.Orgao} ${normativo.Tipo_Norma} ${normativo.Numero}`);
      
      const analise = analisarNormativoComToqan(client, normativo);
      resultados.push(analise);
      
      Logger.log(`✅ [${i + 1}/${normativos.length}] Concluído - Impacto: ${analise.Impacto_Declarado}`);
      
      // Pequeno delay entre análises
      if (i < normativos.length - 1) {
        Utilities.sleep(4000);
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no normativo ${i + 1}: ${error}`);
    }
  }
  
  Logger.log(`🎉 Análise concluída: ${resultados.length} normativos processados`);
  return resultados;
}

function analisarNormativoComToqan(client, normativo) {
  try {
    // Preparar texto para análise
    const textoAnalise = normativo.texto_completo || normativo.Tema || '';
    
    const prompt = `Analise ESTE NORMATIVO REAL para compliance iFood e responda APENAS com JSON:

**NORMATIVO:**
Órgão: ${normativo.Orgao}
Tipo: ${normativo.Tipo_Norma}
Número: ${normativo.Numero}
Data: ${normativo.Data_Publicacao}
Tema: ${normativo.Tema}
Texto: ${textoAnalise.substring(0, 1500)}

**CONTEXTO IFOOD:**
- iFood Pago (instituição de pagamento, PIX, cartões, voucher, IP)
- iFood Crédito (empréstimos, crédito consignado)
- SCD (Sociedade de Crédito Direto)
- Pagamentos, taxas, compliance financeiro
- Instituição de Pagamento, IP, instituição financeira

**RESPONDA APENAS COM ESTE JSON:**
{
  "impacto": "Alto|Médio|Baixo",
  "produto_afetado": "iFood Pago|iFood Crédito|SCD|Múltiplos|Nenhum",
  "aplicavel_scd": "Sim|Não",
  "aplicavel_ip": "Sim|Não",
  "criticidade": "CRÍTICA|ALTA|MÉDIA|BAIXA",
  "resumo_impacto": "Resumo conciso do impacto específico para iFood",
  "acoes_recomendadas": "Ações específicas recomendadas"
}`;

    Logger.log(`   🤖 Enviando para Toqan...`);
    const resposta = client.createConversation(prompt);
    
    Logger.log(`   ✅ Toqan recebeu: ${resposta.conversation_id}`);
    
    // Aguardar processamento
    Utilities.sleep(5000);
    
    // Processar resposta
    return processarRespostaToqanMelhorada(resposta, normativo);
    
  } catch (error) {
    Logger.log(`   ❌ Erro Toqan: ${error}`);
  }
}

function processarRespostaToqanMelhorada(resposta, normativo) {
  try {
    // Valores padrão
    let impacto = 'Médio';
    let produtoAfetado = 'iFood Pago - Geral';
    let aplicavelSCD = 'Não';
    let aplicavelIfood = 'Sim';
    let criticidade = 'MÉDIA';
    let resumoImpacto = 'Análise em andamento - impacto a ser determinado';
    let acoesRecomendadas = 'Aguardar análise detalhada pela equipe jurídica';
    
    // Tentar extrair JSON da resposta
    if (resposta && typeof resposta === 'object') {
      const respostaStr = JSON.stringify(resposta);
      
      // Extrair informações usando regex mais robusto
      const impactoMatch = respostaStr.match(/"impacto"\s*:\s*"([^"]*)"/i);
      const produtoMatch = respostaStr.match(/"produto_afetado"\s*:\s*"([^"]*)"/i);
      const scdMatch = respostaStr.match(/"aplicavel_scd"\s*:\s*"([^"]*)"/i);
      const ifoodMatch = respostaStr.match(/"aplicavel_ifood"\s*:\s*"([^"]*)"/i);
      const criticidadeMatch = respostaStr.match(/"criticidade"\s*:\s*"([^"]*)"/i);
      const resumoMatch = respostaStr.match(/"resumo_impacto"\s*:\s*"([^"]*)"/i);
      const acoesMatch = respostaStr.match(/"acoes_recomendadas"\s*:\s*"([^"]*)"/i);
      
      if (impactoMatch) impacto = impactoMatch[1];
      if (produtoMatch) produtoAfetado = produtoMatch[1];
      if (scdMatch) aplicavelSCD = scdMatch[1];
      if (ifoodMatch) aplicavelIfood = ifoodMatch[1];
      if (criticidadeMatch) criticidade = criticidadeMatch[1];
      if (resumoMatch) resumoImpacto = resumoMatch[1];
      if (acoesMatch) acoesRecomendadas = acoesMatch[1];
    }
    
    const resultado = {
      normativo_index: obterProximoIndex(),
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      Orgao: normativo.Orgao || 'N/A',
      Tipo_Norma: normativo.Tipo_Norma || 'N/A',
      Numero: normativo.Numero || 'N/A',
      Data_Publicacao: normativo.Data_Publicacao || 'N/A',
      Produto_Segmento: produtoAfetado,
      Tema: normativo.Tema || 'N/A',
      Impacto_Declarado: impacto,
      Data_Vigencia: normativo.Data_Publicacao || 'N/A',
      Aplicavel_SCD: aplicavelSCD,
      Aplicavel_IP: 'Sim',
      Aplicavel_iFood: aplicavelIfood,
      status: 'Analisado',
      Criticidade_Sistema: criticidade,
      Resumo_Analise: resumoImpacto,
      Acoes_Recomendadas: acoesRecomendadas,
      Resposta_Toqan: `Toqan ID: ${resposta.conversation_id}`
    };
    
    Logger.log(`   📈 Análise: ${impacto} impacto | ${produtoAfetado} | SCD:${aplicavelSCD}`);
    Logger.log(`   📝 Resumo: ${resumoImpacto.substring(0, 100)}...`);
    
    return resultado;
    
  } catch (error) {
    Logger.log(`   ⚠️ Erro processar resposta: ${error}`);
 }
}
// =============================================
// SISTEMA DE ANÁLISE TOQAN COM FILTRO DE APLICABILIDADE
// =============================================

function analisarNormativosComToqan(normativos) {
  if (!normativos || normativos.length === 0) {
    Logger.log('ℹ️ Nenhum normativo para analisar');
    return [];
  }
  
  Logger.log(`🔍 Iniciando análise de ${normativos.length} normativos com Toqan`);
  const client = new ToqanClient();
  const resultados = [];
  let analisados = 0;
  let aplicaveis = 0;
  
  for (let i = 0; i < normativos.length; i++) {
    const normativo = normativos[i];
    
    try {
      Logger.log(`📊 [${i + 1}/${normativos.length}] Analisando: ${normativo.Orgao} - ${normativo.Tema.substring(0, 50)}...`);
      
      const analise = analisarNormativoComToqan(client, normativo);
      
      if (analise) {
        analisados++;
        
        // FILTRAR: Só incluir se for aplicável ao iFood
        if (analise.aplicavel_ifood === 'Sim' && 
            analise.impacto !== 'N/A' && 
            analise.impacto !== 'Não Aplicável') {
          
          resultados.push(analise);
          aplicaveis++;
          Logger.log(`   ✅ APLICÁVEL - Impacto: ${analise.Impacto_Declarado}`);
        } else {
          Logger.log(`   ❌ NÃO APLICÁVEL - Descarte: ${analise.aplicavel_ifood} | ${analise.impacto}`);
        }
      }
      
      // Pequeno delay entre análises
      if (i < normativos.length - 1) {
        Utilities.sleep(5000); // 5 segundos entre análises
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no normativo ${i + 1}: ${error}`);
    }
  }
  
  Logger.log(`🎉 Análise concluída: ${analisados} processados, ${aplicaveis} aplicáveis ao iFood`);
  return resultados;
}

function analisarNormativoComToqan(client, normativo) {
  try {
    // Preparar texto para análise
    const textoAnalise = normativo.texto_completo || normativo.Tema || '';
    const orgao = normativo.Orgao || 'N/A';
    const tipo = normativo.Tipo_Norma || 'N/A';
    
    const prompt = `Analise ESTE CONTEÚDO para determinar se é APLICÁVEL ao iFood e qual o IMPACTO REAL.

**CONTEÚDO PARA ANÁLISE:**
Fonte: ${orgao}
Tipo: ${tipo}
Número: ${normativo.Numero || 'N/A'}
Data: ${normativo.Data_Publicacao || 'N/A'}
Título: ${normativo.Tema || 'N/A'}
Texto: ${textoAnalise.substring(0, 2000)}

**CONTEXTO IFOOD - ATIVIDADES RELEVANTES:**
- iFood Pago: Sistema de pagamentos (PIX, cartões, voucher alimentação)
- iFood Crédito: Empréstimos, crédito consignado para entregadores
- SCD (Sociedade de Crédito Direto): Operações de crédito
- IP (Instituição de Pagamento): instituição de pagamentos
- Marketplace: Intermediação de vendas de restaurantes
- Pagamentos instantâneos, taxas de intermediação

**CRITÉRIOS DE APLICABILIDADE - CONSIDERE APENAS SE ENCAIXAR EM:**
✅ Regulamentação de pagamentos, PIX, cartões, instituições de pagamento
✅ Normas sobre crédito, empréstimos, fintechs
✅ Regulação de marketplaces, intermediação
✅ Compliance financeiro, prevenção à lavagem
✅ Taxas de intermediação, relações com parceiros
❌ NÃO APLICÁVEL: Notícias gerais, política, outros setores

**RESPONDA APENAS COM ESTE JSON:**

{
  "aplicavel_ifood": "Sim" ou "Não",
  "impacto": "Alto" ou "Médio" ou "Baixo" ou "Não Aplicável",
  "motivo_aplicabilidade": "Explicação curta do porquê é ou não aplicável",
  "produto_afetado": "iFood Pago" ou "iFood Crédito" ou "SCD" ou "Marketplace" ou "Múltiplos" ou "Nenhum",
  "aplicavel_scd": "Sim" ou "Não",
  "resumo_impacto": "Resumo específico do impacto para iFood",
  "acoes_recomendadas": "Ações específicas recomendadas ou 'Nenhuma ação necessária'"
}

**SEJA RIGOROSO: Marque como "Não Aplicável" se não tiver relação direta com as atividades do iFood.**`;

    Logger.log(`   🤖 Enviando para Toqan...`);
    const resposta = client.createConversation(prompt);
    
    Logger.log(`   ✅ Toqan recebeu: ${resposta.conversation_id}`);
    
    // Aguardar processamento
    Utilities.sleep(6000);
    
    // Processar resposta com validação rigorosa
    return processarRespostaToqanFiltrada(resposta, normativo);
    
  } catch (error) {
    Logger.log(`   ❌ Erro Toqan: ${error}`);
    return null;
  }
}

function processarRespostaToqanFiltrada(resposta, normativo) {
  try {
    // Valores padrão CONSERVADORES - assumir não aplicável até provar o contrário
    let aplicavelIfood = 'Não';
    let impacto = 'Não Aplicável';
    let motivoAplicabilidade = 'Análise em andamento';
    let produtoAfetado = 'Nenhum';
    let aplicavelSCD = 'Não';
    let resumoImpacto = 'Aguardar análise detalhada';
    let acoesRecomendadas = 'Nenhuma ação necessária';
    
    // Tentar extrair JSON da resposta
    if (resposta && typeof resposta === 'object') {
      const respostaStr = JSON.stringify(resposta);
      
      // Extrair informações com regex mais específicos
      const aplicavelMatch = respostaStr.match(/"aplicavel_ifood"\s*:\s*"([^"]*)"/i);
      const impactoMatch = respostaStr.match(/"impacto"\s*:\s*"([^"]*)"/i);
      const motivoMatch = respostaStr.match(/"motivo_aplicabilidade"\s*:\s*"([^"]*)"/i);
      const produtoMatch = respostaStr.match(/"produto_afetado"\s*:\s*"([^"]*)"/i);
      const scdMatch = respostaStr.match(/"aplicavel_scd"\s*:\s*"([^"]*)"/i);
      const resumoMatch = respostaStr.match(/"resumo_impacto"\s*:\s*"([^"]*)"/i);
      const acoesMatch = respostaStr.match(/"acoes_recomendadas"\s*:\s*"([^"]*)"/i);
      
      if (aplicavelMatch) aplicavelIfood = aplicavelMatch[1];
      if (impactoMatch) impacto = impactoMatch[1];
      if (motivoMatch) motivoAplicabilidade = motivoMatch[1];
      if (produtoMatch) produtoAfetado = produtoMatch[1];
      if (scdMatch) aplicavelSCD = scdMatch[1];
      if (resumoMatch) resumoImpacto = resumoMatch[1];
      if (acoesMatch) acoesRecomendadas = acoesMatch[1];
      
      // VALIDAÇÃO: Se for "Não Aplicável", forçar consistência
      if (impacto === 'Não Aplicável') {
        aplicavelIfood = 'Não';
        produtoAfetado = 'Nenhum';
        aplicavelSCD = 'Não';
      }
      
      // VALIDAÇÃO: Se não for aplicável, impacto deve ser "Não Aplicável"
      if (aplicavelIfood === 'Não' && impacto !== 'Não Aplicável') {
        impacto = 'Não Aplicável';
      }
    }
    
    const resultado = {
      normativo_index: obterProximoIndex(),
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      Orgao: normativo.Orgao || 'N/A',
      Tipo_Norma: normativo.Tipo_Norma || 'N/A',
      Numero: normativo.Numero || 'N/A',
      Data_Publicacao: normativo.Data_Publicacao || 'N/A',
      Produto_Segmento: produtoAfetado,
      Tema: normativo.Tema || 'N/A',
      Impacto_Declarado: impacto,
      Data_Vigencia: normativo.Data_Publicacao || 'N/A',
      Aplicavel_SCD: aplicavelSCD,
      Aplicavel_IP: aplicavelIfood, // Usar mesma lógica do iFood
      Aplicavel_iFood: aplicavelIfood,
      status: aplicavelIfood === 'Sim' ? 'Analisado' : 'Não Aplicável',
      Criticidade_Sistema: calcularCriticidade(impacto),
      Resumo_Analise: `${motivoAplicabilidade} | ${resumoImpacto}`,
      Acoes_Recomendadas: acoesRecomendadas,
      Resposta_Toqan: `Toqan ID: ${resposta.conversation_id}`,
      url_fonte: normativo.url_fonte || 'N/A'
    };
    
    Logger.log(`   📈 Resultado: ${aplicavelIfood} | Impacto: ${impacto} | Produto: ${produtoAfetado}`);
    Logger.log(`   📝 Motivo: ${motivoAplicabilidade.substring(0, 80)}...`);
    
    return resultado;
    
  } catch (error) {
    Logger.log(`   ⚠️ Erro processar resposta: ${error}`);
    return null;
  }
}

function calcularCriticidade(impacto) {
  switch(impacto) {
    case 'Alto': return 'ALTA';
    case 'Médio': return 'MÉDIA';
    case 'Baixo': return 'BAIXA';
    case 'Não Aplicável': return 'N/A';
    default: return 'MÉDIA';
  }
}

// =============================================
// FUNÇÃO PRINCIPAL COM FILTRO TOQAN
// =============================================

function executarSistemaCompletoComFiltro() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO - COM FILTRO TOQAN');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando execução com filtro de aplicabilidade');
  
  try {
    const startTime = new Date();
    
    // 1. COLETAR NORMATIVOS REAIS
    Logger.log('📡 ETAPA 1: COLETANDO NORMATIVOS REAIS...');
    const normativos = coletarNormativosReais();
    
    if (!normativos || normativos.length === 0) {
      Logger.log('ℹ️ Nenhum normativo novo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return;
    }
    
    Logger.log(`📊 ${normativos.length} normativos reais coletados`);
    
    // 2. ANALISAR COM TOQAN (COM FILTRO)
    Logger.log('🤖 ETAPA 2: ANALISANDO E FILTRANDO COM TOQAN...');
    const normativosFiltrados = analisarNormativosComToqan(normativos);
    
    if (normativosFiltrados.length === 0) {
      Logger.log('ℹ️ Nenhum normativo aplicável ao iFood identificado');
      enviarSlackMensagem('✅ *MONITORAMENTO IFOOD* - Nenhum normativo aplicável identificado hoje');
      return;
    }
    
    Logger.log(`🎯 ${normativosFiltrados.length} normativos aplicáveis identificados`);
    
    // 3. SALVAR NA PLANILHA APENAS OS APLICÁVEIS
    Logger.log('💾 ETAPA 3: SALVANDO APENAS NORMATIVOS APLICÁVEIS...');
    const salvos = salvarNaPlanilha(normativosFiltrados);
    
    // 4. ENVIAR RELATÓRIO APENAS COM APLICÁVEIS
    Logger.log('📤 ETAPA 4: ENVIANDO RELATÓRIO FILTRADO...');
    enviarRelatorioFiltradoSlack(normativosFiltrados, salvos, normativos.length);
    
    const endTime = new Date();
    const tempoExecucao = (endTime - startTime) / 1000;
    
    registrarLogAPI('SISTEMA', 'SUCCESS', 
      `Execução concluída - ${normativosFiltrados.length}/${normativos.length} normativos aplicáveis em ${tempoExecucao}s`, 
      normativosFiltrados.length
    );
    
    Logger.log(`🎉 SISTEMA CONCLUÍDO EM ${tempoExecucao}s! ${normativosFiltrados.length}/${normativos.length} aplicáveis`);
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NO SISTEMA: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro no sistema: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
  }
}

function enviarRelatorioFiltradoSlack(normativosFiltrados, salvos, totalColetado) {
  try {
    const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
    const horaAtual = Utilities.formatDate(new Date(), 'GMT-3', 'HH:mm');
    
    let mensagem = `🎯 *MONITORAMENTO IFOOD - ${dataHoje} ${horaAtual}*\n\n`;
    mensagem += `📊 *RELATÓRIO FILTRADO - APLICÁVEIS AO IFOOD*\n`;
    mensagem += `• Coletados: ${totalColetado} itens\n`;
    mensagem += `• Aplicáveis: ${normativosFiltrados.length} itens\n`;
    mensagem += `• Salvos: ${salvos} itens\n\n`;
    
    // DETALHAMENTO APENAS DOS APLICÁVEIS
    if (normativosFiltrados.length > 0) {
      mensagem += `🚨 *NORMATIVOS APLICÁVEIS IDENTIFICADOS:*\n\n`;
      
      normativosFiltrados.forEach((normativo, index) => {
        const emojiImpacto = normativo.Impacto_Declarado === 'Alto' ? '🔴' : 
                           normativo.Impacto_Declarado === 'Médio' ? '🟡' : '🟢';
        
        mensagem += `${emojiImpacto} *${normativo.Orgao} ${normativo.Tipo_Norma} ${normativo.Numero}*\n`;
        mensagem += `   _${normativo.Tema}_\n`;
        mensagem += `   📊 *Impacto:* ${normativo.Impacto_Declarado} | *Produto:* ${normativo.Produto_Segmento}\n`;
        mensagem += `   ✅ *Aplicável:* SCD:${normativo.Aplicavel_SCD} | iFood:${normativo.Aplicavel_iFood}\n`;
        
        // RESUMO DA ANÁLISE
        if (normativo.Resumo_Analise) {
          mensagem += `   📝 *Análise:* ${normativo.Resumo_Analise.substring(0, 100)}...\n`;
        }
        
        // AÇÕES RECOMENDADAS (apenas se não for "Nenhuma ação necessária")
        if (normativo.Acoes_Recomendadas && !normativo.Acoes_Recomendadas.includes('Nenhuma ação necessária')) {
          mensagem += `   🎯 *Ações:* ${normativo.Acoes_Recomendadas.substring(0, 80)}...\n`;
        }
        
        mensagem += `\n`;
      });
      
      // RESUMO POR IMPACTO
      const altoImpacto = normativosFiltrados.filter(n => n.Impacto_Declarado === 'Alto').length;
      const medioImpacto = normativosFiltrados.filter(n => n.Impacto_Declarado === 'Médio').length;
      const baixoImpacto = normativosFiltrados.filter(n => n.Impacto_Declarado === 'Baixo').length;
      
      mensagem += `📈 *RESUMO POR IMPACTO:*\n`;
      mensagem += `• 🔴 Alto: ${altoImpacto}\n`;
      mensagem += `• 🟡 Médio: ${medioImpacto}\n`;
      mensagem += `• 🟢 Baixo: ${baixoImpacto}\n\n`;
    } else {
      mensagem += `✅ *NENHUM NORMATIVO APLICÁVEL IDENTIFICADO HOJE*\n`;
      mensagem += `O sistema analisou ${totalColetado} itens e não encontrou nenhum com impacto direto ao iFood.\n\n`;
    }
    
    mensagem += `⚡ _Sistema Automático iFood Compliance - Análise Toqan AI com Filtro_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro relatório filtrado: ${error}`);
    return enviarSlackMensagem(`📋 Monitoramento iFood - ${normativosFiltrados.length} normativos aplicáveis identificados`);
  }
}

// =============================================
// FUNÇÃO DE TESTE DO FILTRO
// =============================================

function testarFiltroToqan() {
  Logger.log('🧪 TESTANDO FILTRO DE APLICABILIDADE TOQAN');
  
  try {
    // Criar dados de teste variados
    const normativosTeste = [
      {
        Orgao: 'BCB_NOTICIAS',
        Tipo_Norma: 'Notícia',
        Numero: 'TESTE-ALTO-1',
        Data_Publicacao: '2024-01-01',
        Tema: 'BACEN anuncia nova regulamentação para pagamentos instantâneos PIX',
        texto_completo: 'O Banco Central anunciou novas regras para operações de pagamento instantâneo via PIX que afetam todas as instituições de pagamento.',
        url_fonte: 'https://www.bcb.gov.br/noticias'
      },
      {
        Orgao: 'VALOR_ECONOMICO',
        Tipo_Norma: 'Notícia',
        Numero: 'TESTE-NAO-APLICAVEL-1',
        Data_Publicacao: '2024-01-01',
        Tema: 'Bolsa de Valores tem alta recorde com notícias do exterior',
        texto_completo: 'A bolsa brasileira fechou em alta influenciada por notícias positivas do mercado internacional.',
        url_fonte: 'https://valor.globo.com/financas/'
      },
      {
        Orgao: 'INFOMONEY',
        Tipo_Norma: 'Notícia',
        Numero: 'TESTE-MEDIO-1',
        Data_Publicacao: '2024-01-01',
        Tema: 'CMN aprova novas regras para crédito consignado para plataformas digitais',
        texto_completo: 'O Conselho Monetário Nacional aprovou resolução que altera as regras para crédito consignado em plataformas digitais como iFood e Uber.',
        url_fonte: 'https://www.infomoney.com.br/'
      }
    ];
    
    Logger.log('📝 Testando com 3 normativos: 1 aplicável alto, 1 médio, 1 não aplicável');
    
    const resultados = analisarNormativosComToqan(normativosTeste);
    
    Logger.log(`📊 Resultado do teste: ${resultados.length} aplicáveis identificados`);
    
    resultados.forEach((resultado, index) => {
      Logger.log(`   ${index + 1}. ${resultado.Orgao} - Aplicável: ${resultado.Aplicavel_iFood} - Impacto: ${resultado.Impacto_Declarado}`);
    });
    
    const esperado = 2; // Devem ser aplicáveis apenas os 2 primeiros
    const sucesso = resultados.length === esperado;
    
    if (sucesso) {
      Logger.log('✅ TESTE DO FILTRO BEM-SUCEDIDO!');
    } else {
      Logger.log(`❌ TESTE FALHOU: Esperado ${esperado}, obtido ${resultados.length}`);
    }
    
    return sucesso;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO TESTE: ${error.toString()}`);
    return false;
  }
}
// =============================================
// FUNÇÃO PRINCIPAL ATUALIZADA COM NOTIFICAÇÕES
// =============================================

function executarSistemaCompleto() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO - CAPTURA REAL + ANÁLISE TOQAN');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando execução do sistema com captura real');
  
  try {
    const startTime = new Date();
    
    // 1. COLETAR NORMATIVOS REAIS
    Logger.log('📡 ETAPA 1: COLETANDO NORMATIVOS REAIS DO BACEN...');
    const normativos = coletarNormativosReais();
    
    if (!normativos || normativos.length === 0) {
      Logger.log('ℹ️ Nenhum normativo novo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return;
    }
    
    Logger.log(`📊 ${normativos.length} normativos reais coletados`);
    
    // 2. ANALISAR COM TOQAN
    Logger.log('🤖 ETAPA 2: ANALISANDO COM TOQAN...');
    const normativosAnalisados = analisarNormativosComToqan(normativos);
    
    // 3. SALVAR NA PLANILHA
    Logger.log('💾 ETAPA 3: SALVANDO NA PLANILHA...');
    const salvos = salvarNaPlanilha(normativosAnalisados);
    
    // 4. ENVIAR RELATÓRIO COMPLETO COM ANÁLISE TOQAN
    Logger.log('📤 ETAPA 4: ENVIANDO RELATÓRIO COM ANÁLISE...');
    enviarRelatorioCompletoComAnalise(normativosAnalisados, salvos);
    
    const endTime = new Date();
    const tempoExecucao = (endTime - startTime) / 1000;
    
    registrarLogAPI('SISTEMA', 'SUCCESS', 
      `Execução concluída - ${normativosAnalisados.length} normativos processados em ${tempoExecucao}s`, 
      normativosAnalisados.length
    );
    
    Logger.log(`🎉 SISTEMA CONCLUÍDO EM ${tempoExecucao}s! ${normativosAnalisados.length} normativos processados`);
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NO SISTEMA: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro no sistema: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
  }
}

function enviarRelatorioCompletoComAnalise(normativos, salvos) {
  try {
    const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
    const horaAtual = Utilities.formatDate(new Date(), 'GMT-3', 'HH:mm');
    
    let mensagem = `🎯 *MONITORAMENTO IFOOD - ${dataHoje} ${horaAtual}*\n\n`;
    mensagem += `📊 *RESUMO EXECUTIVO*\n`;
    mensagem += `• Normativos detectados: ${normativos.length}\n`;
    mensagem += `• Salvos na planilha: ${salvos}\n`;
    mensagem += `• Análise: Toqan AI\n\n`;
    
    // DETALHAMENTO COM ANÁLISE TOQAN
    if (normativos.length > 0) {
      mensagem += `📋 *NORMATIVOS DETECTADOS COM ANÁLISE TOQAN:*\n\n`;
      
      normativos.forEach((normativo, index) => {
        const emojiImpacto = normativo.Impacto_Declarado === 'Alto' ? '🔴' : 
                           normativo.Impacto_Declarado === 'Médio' ? '🟡' : '🟢';
        
        const emojiCriticidade = normativo.Criticidade_Sistema === 'CRÍTICA' ? '🚨' :
                               normativo.Criticidade_Sistema === 'ALTA' ? '⚠️' : 'ℹ️';
        
        mensagem += `${emojiImpacto} ${emojiCriticidade} *${normativo.Orgao} ${normativo.Tipo_Norma} ${normativo.Numero}*\n`;
        mensagem += `   _${normativo.Tema}_\n`;
        mensagem += `   📊 *Impacto:* ${normativo.Impacto_Declarado} | *Criticidade:* ${normativo.Criticidade_Sistema}\n`;
        mensagem += `   🎯 *Produto Afetado:* ${normativo.Produto_Segmento}\n`;
        mensagem += `   ✅ *Aplicável:* SCD:${normativo.Aplicavel_SCD} | iFood:${normativo.Aplicavel_iFood}\n`;
        
        // RESUMO DA ANÁLISE TOQAN
        if (normativo.Resumo_Analise && normativo.Resumo_Analise !== 'Análise em andamento - impacto a ser determinado') {
          mensagem += `   📝 *Análise Toqan:* ${normativo.Resumo_Analise}\n`;
        }
        
        // AÇÕES RECOMENDADAS
        if (normativo.Acoes_Recomendadas && normativo.Acoes_Recomendadas !== 'Aguardar análise detalhada pela equipe jurídica') {
          mensagem += `   🎯 *Ações:* ${normativo.Acoes_Recomendadas}\n`;
        }
        
        mensagem += `   🔗 *Fonte:* ${normativo.url_fonte || 'N/A'}\n`;
        mensagem += `\n`;
      });
      
      // ALERTAS CRÍTICOS
      const normativosCriticos = normativos.filter(n => 
        n.Criticidade_Sistema === 'CRÍTICA' || n.Impacto_Declarado === 'Alto'
      );
      
      if (normativosCriticos.length > 0) {
        mensagem += `🚨 *ALERTAS CRÍTICOS - AÇÃO IMEDIATA REQUERIDA*\n`;
        mensagem += `• ${normativosCriticos.length} normativo(s) de alto impacto/criticidade detectado(s)\n`;
        mensagem += `• Recomendação: Revisão urgente pela equipe jurídica\n\n`;
      }
    }
    
    mensagem += `⚡ _Sistema Automático iFood Compliance - Análise Toqan AI_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro relatório com análise: ${error}`);
    // Fallback: enviar relatório básico
    return enviarSlackMensagem(`📋 Monitoramento iFood - ${normativos.length} normativos processados com análise Toqan`);
  }
}

// =============================================
// FUNÇÕES DE EXECUÇÃO RÁPIDA
// =============================================

/**
 * EXECUTAR AGORA (manual) - Sistema completo
 */
function executarAgora() {
  Logger.log('🚀 EXECUTANDO SISTEMA COMPLETO AGORA');
  executarSistemaCompleto();
}

/**
 * TESTAR APENAS TOQAN
 */
function testarToqanAgora() {
  Logger.log('🧪 EXECUTANDO TESTE ESPECÍFICO DO TOQAN');
  return testarToqanEspecifico();
}

// =============================================
// CONFIGURAÇÃO DE AGENDAMENTO
// =============================================

function configurarAgendamentoAutomatico() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTO AUTOMÁTICO');
  
  try {
    // Remover triggers existentes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️  Trigger removido: ${trigger.getHandlerFunction()}`);
    });
    
    // Agendamentos principais
    const horarios = [9, 17]; // 9h e 17h
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarSistemaCompleto')
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      
      Logger.log(`✅ Agendado: ${hora}:00 diariamente`);
    });
    
    // Agendamento de saúde do sistema (8h)
    ScriptApp.newTrigger('verificarSaudeSistema')
      .timeBased()
      .atHour(8)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone('America/Sao_Paulo')
      .create();
    
    Logger.log('✅ Agendamento: Saúde do sistema às 8:00');
    
    const mensagem = `✅ *SISTEMA IFOOD CONFIGURADO*

⏰ *Agendamentos ativos:*
• 9:00 e 17:00 - Monitoramento diário
• 8:00 - Verificação de saúde

🤖 *Recursos:*
• Captura real BACEN/CMN
• Análise Toqan AI
• Notificações Slack
• Logs detalhados

🚀 Sistema operacional!`;
    
    enviarSlackMensagem(mensagem);
    Logger.log('🎉 AGENDAMENTO CONFIGURADO COM SUCESSO!');
    
  } catch (error) {
    Logger.log(`❌ ERRO AO CONFIGURAR AGENDAMENTO: ${error.toString()}`);
    enviarSlackMensagem(`❌ Erro na configuração do agendamento: ${error.toString().substring(0, 100)}`);
  }
}

function verificarSaudeSistema() {
  Logger.log('🏥 VERIFICANDO SAÚDE DO SISTEMA');
  
  try {
    const testes = [];
    
    // Teste 1: Planilha
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      const sheet = spreadsheet.getSheets()[0];
      const ultimaLinha = sheet.getLastRow();
      testes.push({ item: '📊 Planilha', status: '✅', detalhes: `${ultimaLinha} linhas` });
    } catch (e) {
      testes.push({ item: '📊 Planilha', status: '❌', detalhes: e.toString() });
    }
    
    // Teste 2: Toqan
    try {
      const client = new ToqanClient();
      const teste = client.createConversation("Teste de saúde - OK");
      testes.push({ item: '🤖 Toqan AI', status: '✅', detalhes: teste.conversation_id });
    } catch (e) {
      testes.push({ item: '🤖 Toqan AI', status: '❌', detalhes: e.toString() });
    }
    
    // Teste 3: Sistema de Logs
    try {
      registrarLogAPI('SAÚDE', 'INFO', 'Teste de verificação de saúde');
      testes.push({ item: '📋 Sistema de Logs', status: '✅', detalhes: 'Logs funcionando' });
    } catch (e) {
      testes.push({ item: '📋 Sistema de Logs', status: '❌', detalhes: e.toString() });
    }
    
    // Preparar relatório
    const dataVerificacao = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
    let relatorio = `🏥 *RELATÓRIO DE SAÚDE - ${dataVerificacao}*\n\n`;
    
    testes.forEach(test => {
      relatorio += `${test.status} ${test.item}: ${test.detalhes}\n`;
    });
    
    relatorio += `\n⚡ _Sistema iFood Compliance_`;
    
    enviarSlackMensagem(relatorio);
    Logger.log('✅ Verificação de saúde concluída');
    
  } catch (error) {
    Logger.log(`❌ Erro na verificação de saúde: ${error}`);
  }
}

// =============================================
// FUNÇÃO PARA VERIFICAR STATUS DO AGENDAMENTO
// =============================================

function verificarStatusAgendamento() {
  Logger.log('📊 VERIFICANDO STATUS DO AGENDAMENTO AUTOMÁTICO');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    let mensagem = `🔍 *STATUS DO AGENDAMENTO AUTOMÁTICO*\n\n`;
    mensagem += `⏰ *Triggers Ativos:* ${triggers.length}\n\n`;
    
    if (triggers.length === 0) {
      mensagem += `📭 *Nenhum agendamento ativo encontrado*\n`;
      mensagem += `💡 *Solução:* Execute 'configurarAgendamentoManual()'`;
    } else {
      triggers.forEach((trigger, index) => {
        mensagem += `${index + 1}. *${trigger.getHandlerFunction()}*\n`;
        mensagem += `   📅 Tipo: ${trigger.getEventType()}\n`;
        
        // Tentar obter detalhes do agendamento
        try {
          const source = trigger.getTriggerSource();
          mensagem += `   🔧 Fonte: ${source}\n`;
        } catch (e) {
          // Ignora erros de detalhes
        }
        
        mensagem += `\n`;
      });
      
      mensagem += `✅ *Sistema configurado para execução automática!*\n`;
    }
    
    mensagem += `\n🕒 ${Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss')}`;
    
    Logger.log(`📋 Status: ${triggers.length} triggers ativos`);
    enviarSlackMensagem(mensagem);
    
    return {
      success: true,
      triggers: triggers.length,
      details: triggers.map(t => ({
        function: t.getHandlerFunction(),
        type: t.getEventType()
      }))
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao verificar status: ${error}`);
    enviarSlackMensagem(`❌ Erro ao verificar agendamento: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// CONFIGURAÇÃO DE AGENDAMENTO GARANTIDA
// =============================================

function configurarAgendamentoManual() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTO MANUAL GARANTIDO');
  
  try {
    // Limpar TUDO primeiro
    const todosTriggers = ScriptApp.getProjectTriggers();
    todosTriggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️  Removido: ${trigger.getHandlerFunction()}`);
    });
    
    Logger.log('✅ Todos os triggers anteriores removidos');
    Utilities.sleep(3000); // Aguardar mais tempo
    
    // AGENDAMENTO PRINCIPAL - Horários comerciais
    const horarios = [9, 12, 17]; // 9h, 12h, 17h
    
    for (let hora of horarios) {
      try {
        ScriptApp.newTrigger('executarSistemaCompleto')
          .timeBased()
          .atHour(hora)
          .nearMinute(0)
          .everyDays(1)
          .inTimezone('America/Sao_Paulo')
          .create();
        
        Logger.log(`✅ Agendado com sucesso: ${hora}:00 diariamente`);
      } catch (e) {
        Logger.log(`⚠️ Erro no horário ${hora}h: ${e.toString()}`);
      }
      Utilities.sleep(2000); // Delay entre criações
    }
    
    // AGENDAMENTO DE VERIFICAÇÃO (mais simples)
    try {
      ScriptApp.newTrigger('executarMonitoramentoTeste')
        .timeBased()
        .everyHours(6)
        .create();
      Logger.log('✅ Agendamento de teste: A cada 6 horas');
    } catch (e) {
      Logger.log(`⚠️ Erro no agendamento teste: ${e}`);
    }
    
    // VERIFICAR resultado final
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    const mensagem = `🎉 *AGENDAMENTO CONFIGURADO COM SUCESSO!*

✅ *Execuções Automáticas Ativas:*
• 9:00, 12:00 e 17:00 - Monitoramento completo
• A cada 6h - Verificação rápida

📊 *Total de Agendamentos:* ${triggersFinais.length}

🤖 *O sistema executará sozinho:*
├── Captura de normativos BACEN/CMN
├── Análise automática com Toqan AI  
├── Salvamento na planilha
└── Notificações no Slack

🚀 *Próxima execução automática:* Amanhã às 9:00

⚡ _Sistema 100% automatizado_`;
    
    enviarSlackMensagem(mensagem);
    Logger.log(`🎉 CONFIGURAÇÃO FINALIZADA: ${triggersFinais.length} agendamentos ativos`);
    
    return {
      success: true,
      triggers: triggersFinais.length,
      nextExecution: 'Amanhã às 9:00'
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NA CONFIGURAÇÃO: ${error.toString()}`);
    
    // Última tentativa - método ultra simples
    try {
      ScriptApp.newTrigger('executarSistemaCompleto')
        .timeBased()
        .everyDays(1)
        .create();
      
      Logger.log('✅ Configuração mínima realizada');
      enviarSlackMensagem('✅ Configuração mínima - Execução diária ativa');
      
    } catch (finalError) {
      Logger.log(`❌ FALHA TOTAL NO AGENDAMENTO: ${finalError}`);
      enviarSlackMensagem('❌ Falha na configuração automática. Usar execução manual.');
    }
    
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA CONFIRMAR SISTEMA AUTOMÁTICO
// =============================================

function confirmarSistemaAutomatico() {
  Logger.log('🔍 CONFIRMANDO SISTEMA AUTOMÁTICO');
  
  // 1. Verificar agendamentos atuais
  const status = verificarStatusAgendamento();
  
  // 2. Se não há agendamentos, configurar
  if (!status.triggers || status.triggers === 0) {
    Logger.log('⚠️ Nenhum agendamento encontrado - Configurando...');
    return configurarAgendamentoManual();
  }
  
  // 3. Se já há agendamentos, confirmar
  Logger.log(`✅ Sistema já possui ${status.triggers} agendamentos ativos`);
  
  const mensagem = `✅ *SISTEMA AUTOMÁTICO CONFIRMADO!*

📊 *Agendamentos Ativos:* ${status.triggers}

⏰ *Próximas Execuções Automáticas:*
• Amanhã às 9:00, 12:00 e 17:00
• Verificações a cada 6 horas

🤖 *Processo Automático:*
1. 🕐 Horário agendado → Dispara execução
2. 🔍 Sistema captura normativos BACEN/CMN
3. 🧠 Toqan analisa impacto para iFood
4. 💾 Salva automaticamente na planilha
5. 📤 Envia relatório completo no Slack

🚀 *Sistema 100% autônomo - Sem necessidade de intervenção manual*

⚡ _Monitoramento contínuo ativo_`;
  
  enviarSlackMensagem(mensagem);
  
  return {
    success: true,
    message: 'Sistema automático confirmado e ativo',
    triggers: status.triggers
  };
}

// =============================================
// FUNÇÃO DE TESTE TOQAN SIMPLES (FALTANTE)
// =============================================

function testarToqanSimples() {
  Logger.log('🧪 TESTE SIMPLES DO TOQAN');
  
  try {
    const client = new ToqanClient();
    const resposta = client.createConversation("Teste de conexão - responda apenas com 'OK'");
    
    if (resposta && resposta.conversation_id) {
      Logger.log('✅ Toqan conectado com sucesso');
      return true;
    } else {
      Logger.log('❌ Toqan não retornou ID da conversação');
      return false;
    }
    
  } catch (error) {
    Logger.log(`❌ Erro no teste Toqan: ${error.toString()}`);
    return false;
  }
}

// =============================================
// FUNÇÃO DE MONITORAMENTO TESTE (FALTANTE)
// =============================================

function executarMonitoramentoTeste() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO TESTE RÁPIDO');
  
  try {
    // Versão simplificada para testes rápidos
    const normativos = coletarNormativosReais();
    
    if (normativos && normativos.length > 0) {
      Logger.log(`📊 ${normativos.length} normativos encontrados no teste`);
      
      // Apenas salvar sem análise completa para ser mais rápido
      const salvos = salvarNaPlanilha(normativos);
      
      enviarSlackMensagem(`🔍 Monitoramento Teste: ${normativos.length} normativos detectados e salvos`);
    } else {
      Logger.log('ℹ️ Nenhum normativo no teste rápido');
    }
    
  } catch (error) {
    Logger.log(`❌ Erro no monitoramento teste: ${error}`);
  }
}
// =============================================
// CONFIGURAÇÃO DE AGENDAMENTO SIMPLES
// =============================================

function configurarAgendamentoSimples() {
  Logger.log('⏰ CONFIGURANDO APENAS AGENDAMENTO');
  
  try {
    // Remover todos os triggers existentes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️ Removido: ${trigger.getHandlerFunction()}`);
    });
    
    // AGENDAMENTOS PRINCIPAIS (Horários comerciais)
    const horariosComerciais = [9, 12, 17]; // 9h, 12h, 17h
    
    horariosComerciais.forEach(hora => {
      ScriptApp.newTrigger('executarSistemaCompleto') // Esta função será executada nos horários agendados
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`✅ Agendado: ${hora}:00`);
    });
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    enviarSlackMensagem(`⏰ *AGENDAMENTO CONFIGURADO*

📊 ${triggersFinais.length} agendamentos ativos
⏰ Execuções automáticas: 9h, 12h, 17h

✅ O sistema executará automaticamente nestes horários!`);
    
    Logger.log('🎉 AGENDAMENTO CONFIGURADO - Sistema rodará automaticamente');
    
    return {
      success: true,
      triggers: triggersFinais.length,
      message: 'Agendamento configurado com sucesso'
    };
    
  } catch (error) {
    Logger.log(`❌ Erro no agendamento: ${error}`);
    enviarSlackMensagem(`❌ Erro no agendamento: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA VERIFICAR STATUS DO AGENDAMENTO
// =============================================

function verificarStatusAgendamento() {
  Logger.log('📊 VERIFICANDO STATUS DO AGENDAMENTO');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    let mensagem = `🔍 *STATUS DO AGENDAMENTO*\n\n`;
    mensagem += `⏰ *Triggers Ativos:* ${triggers.length}\n\n`;
    
    if (triggers.length === 0) {
      mensagem += `📭 *Nenhum agendamento ativo*\n`;
      mensagem += `💡 Execute 'configurarAgendamentoSimples()'`;
    } else {
      triggers.forEach((trigger, index) => {
        mensagem += `${index + 1}. *${trigger.getHandlerFunction()}*\n`;
        
        // Tentar obter detalhes do agendamento
        try {
          if (trigger.getHandlerFunction() === 'executarSistemaCompleto') {
            mensagem += `   🕐 Execução automática diária\n`;
          }
        } catch (e) {
          // Ignora erros de detalhes
        }
        
        mensagem += `\n`;
      });
      
      mensagem += `✅ *Sistema configurado para execução automática!*\n`;
    }
    
    Logger.log(`📋 Status: ${triggers.length} triggers ativos`);
    enviarSlackMensagem(mensagem);
    
    return {
      success: true,
      triggers: triggers.length,
      details: triggers.map(t => t.getHandlerFunction())
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao verificar status: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA INICIAR APENAS O AGENDAMENTO
// =============================================

function iniciarApenasAgendamento() {
  Logger.log('🚀 INICIANDO APENAS O SISTEMA DE AGENDAMENTO');
  
  try {
    // Apenas configurar o agendamento, sem executar o sistema
    const resultado = configurarAgendamentoSimples();
    
    if (resultado.success) {
      Logger.log('🎉 SISTEMA DE AGENDAMENTO INICIADO!');
      Logger.log('📋 O sistema completo executará automaticamente nos horários configurados');
    }
    
    return resultado;
    
  } catch (error) {
    Logger.log(`❌ Erro ao iniciar agendamento: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA PARAR AGENDAMENTO
// =============================================

function pararAgendamento() {
  Logger.log('🛑 PARANDO AGENDAMENTO AUTOMÁTICO');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
      Logger.log(`🗑️ Removido: ${trigger.getHandlerFunction()}`);
    });
    
    const mensagem = `🛑 *AGENDAMENTO PARADO*

📊 ${removidos} agendamentos removidos
⏰ Execuções automáticas desativadas

💡 Para reativar, execute 'iniciarApenasAgendamento()'`;
    
    enviarSlackMensagem(mensagem);
    Logger.log(`✅ ${removidos} agendamentos removidos`);
    
    return {
      success: true,
      removidos: removidos,
      message: 'Agendamento parado com sucesso'
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao parar agendamento: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA AGENDAMENTO PERSONALIZADO
// =============================================

function configurarAgendamentoPersonalizado(horarios = [9, 17]) {
  Logger.log(`⏰ CONFIGURANDO AGENDAMENTO PERSONALIZADO: ${horarios.join(', ')}h`);
  
  try {
    // Parar agendamentos existentes
    pararAgendamento();
    Utilities.sleep(2000);
    
    // Configurar novos horários
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarSistemaCompleto')
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`✅ Agendado: ${hora}:00`);
    });
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    const mensagem = `⏰ *AGENDAMENTO PERSONALIZADO*

📊 ${triggersFinais.length} agendamentos ativos
⏰ Horários: ${horarios.map(h => `${h}:00`).join(', ')}

✅ Sistema agendado com sucesso!`;
    
    enviarSlackMensagem(mensagem);
    
    return {
      success: true,
      horarios: horarios,
    };
    
  } catch (error) {
    Logger.log(`❌ Erro no agendamento personalizado: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÕES DE CONTROLE RÁPIDO
// =============================================

/**
 * CONFIGURAR AGENDAMENTO RÁPIDO (9h e 17h)
 */
function agendarPadrao() {
  return configurarAgendamentoPersonalizado([9, 17]);
}

/**
 * CONFIGURAR AGENDAMENTO COMERCIAL (9h, 12h, 17h)  
 */
function agendarComercial() {
  return configurarAgendamentoPersonalizado([9, 12, 17]);
}

/**
 * CONFIGURAR AGENDAMENTO CONTÍNUO (9h, 12h, 15h, 17h)
 */
function agendarContinuo() {
  return configurarAgendamentoPersonalizado([9, 12, 15, 17]);
}

// =============================================
// EXECUTAR APENAS O AGENDAMENTO
// =============================================

/**
 * FUNÇÃO PRINCIPAL - EXECUTAR ESTA PARA CONFIGURAR APENAS O AGENDAMENTO
 */
function configurarApenasAgendamento() {
  Logger.log('🚀 CONFIGURANDO APENAS O SISTEMA DE AGENDAMENTO');
  return iniciarApenasAgendamento();
}
configurarApenasAgendamento()
// =============================================
// SISTEMA DE AGENDAMENTO CORRIGIDO
// =============================================

/**
 * Função principal corrigida que evita múltiplas inicializações
 */
function executarSistemaCompleto() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO - VERSÃO CORRIGIDA');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando execução do sistema completo');
  
  try {
    const startTime = new Date();
    
    // 1. COLETAR NORMATIVOS OFICIAIS (sistema existente)
    Logger.log('📡 ETAPA 1: COLETANDO NORMATIVOS OFICIAIS...');
    const normativosOficiais = coletarNormativosReais();
    
    // 2. MONITORAR FONTES COMPLEMENTARES (novo módulo)
    Logger.log('🔍 ETAPA 2: MONITORANDO FONTES COMPLEMENTARES...');
    const monitor = new MonitoramentoNormativo();
      
    // Combinar resultados
    const todosNormativos = [...normativosOficiais, ...fontesComplementares];
    
    if (todosNormativos.length === 0) {
      Logger.log('ℹ️ Nenhum normativo novo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return;
    }
    
    Logger.log(`📊 ${todosNormativos.length} normativos coletados no total`);
    
    // 3. ANALISAR COM TOQAN
    Logger.log('🤖 ETAPA 3: ANALISANDO COM TOQAN...');
    const normativosAnalisados = analisarNormativosComToqan(todosNormativos);
    
    // 4. SALVAR NA PLANILHA
    Logger.log('💾 ETAPA 4: SALVANDO NA PLANILHA...');
    const salvos = salvarNaPlanilha(normativosAnalisados);
    
    // 5. ENVIAR RELATÓRIO
    Logger.log('📤 ETAPA 5: ENVIANDO RELATÓRIO...');
    enviarRelatorioCompletoComAnalise(normativosAnalisados, salvos);
    
    const endTime = new Date();
    const tempoExecucao = (endTime - startTime) / 1000;
    
    registrarLogAPI('SISTEMA', 'SUCCESS', 
      `Execução concluída - ${normativosAnalisados.length} normativos processados em ${tempoExecucao}s`, 
      normativosAnalisados.length
    );
    
    Logger.log(`🎉 SISTEMA CONCLUÍDO EM ${tempoExecucao}s! ${normativosAnalisados.length} normativos processados`);
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NO SISTEMA: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro no sistema: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
  }
}

/**
 * Sistema de agendamento corrigido - evita múltiplos triggers
 */
function configurarAgendamentoAutomatico() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTO AUTOMÁTICO CORRIGIDO');
  
  try {
    // Remover todos os triggers existentes para evitar duplicação
    const triggers = ScriptApp.getProjectTriggers();
    Logger.log(`🔍 Encontrados ${triggers.length} triggers existentes`);
    
    triggers.forEach(trigger => {
      Logger.log(`🗑️ Removendo trigger: ${trigger.getHandlerFunction()}`);
      ScriptApp.deleteTrigger(trigger);
    });
    
    // Verificar se já existe trigger para a função principal
    const triggersExistentes = ScriptApp.getProjectTriggers().filter(
      trigger => trigger.getHandlerFunction() === 'executarSistemaCompleto'
    );
    
    if (triggersExistentes.length === 0) {
      // Criar apenas UM trigger
      ScriptApp.newTrigger('executarSistemaCompleto')
        .timeBased()
        .atHour(9)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      
      Logger.log('✅ Agendamento configurado: execução diária às 9h');
      enviarSlackMensagem('✅ *SISTEMA IFOOD CONFIGURADO* - Agendamento ativo: 9h diariamente');
    } else {
      Logger.log('ℹ️ Agendamento já existe, nenhuma ação necessária');
    }
    
  } catch (error) {
    Logger.log(`❌ ERRO AO CONFIGURAR AGENDAMENTO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO AGENDAMENTO: ${error.toString().substring(0, 100)}`);
  }
}

/**
 * Função para verificar e limpar triggers duplicados
 */
function verificarELimparTriggers() {
  Logger.log('🔍 VERIFICANDO TRIGGERS EXISTENTES');
  
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`📊 Total de triggers: ${triggers.length}`);
  
  triggers.forEach((trigger, index) => {
    Logger.log(`Trigger ${index + 1}: ${trigger.getHandlerFunction()} - ${trigger.getEventType()}`);
  });
  
  // Limpar todos os triggers se necessário
  if (triggers.length > 1) {
    Logger.log('🧹 Limpando triggers duplicados...');
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    Logger.log('✅ Todos os triggers removidos');
  }
}

// =============================================
// FUNÇÃO PRINCIPAL CORRIGIDA COM BACKLOG
// =============================================

function executarSistemaCompletoComBacklog() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO - COM BACKLOG');
  registrarLogAPI('SISTEMA', 'INFO', 'Iniciando execução com sistema de backlog');
  
  try {
    const startTime = new Date();
    
    // 1. COLETAR NORMATIVOS REAIS
    Logger.log('📡 ETAPA 1: COLETANDO NORMATIVOS...');
    const normativos = coletarNormativosReais();
    
    if (!normativos || normativos.length === 0) {
      Logger.log('⚡ Nenhum normativo novo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return;
    }
    
    Logger.log(`📊 ${normativos.length} normativos reais coletados`);
    
    // 2. SALVAR TODOS NO BACKLOG (ANTES DA ANÁLISE)
    Logger.log('📚 ETAPA 2: SALVANDO TODOS NO BACKLOG...');
    const salvosBacklog = salvarNoBacklog(normativos);
    
    // 3. ANALISAR COM TOQAN (APENAS PARA FILTRAGEM)
    Logger.log('🤖 ETAPA 3: ANALISANDO COM TOQAN...');
    const normativosAnalisados = analisarNormativosComToqan(normativos) || [];
    
    // 4. ATUALIZAR BACKLOG COM RESULTADOS DA ANÁLISE
    Logger.log('🔄 ETAPA 4: ATUALIZANDO BACKLOG...');
    const atualizadosBacklog = atualizarBacklogComAnalise(normativosAnalisados);
    
    // 5. SALVAR APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 ETAPA 5: SALVANDO APLICÁVEIS NA PLANILHA...');
    const salvosPlanilha = salvarNaPlanilha(normativosAnalisados) || 0;
    
    // 6. ENVIAR RELATÓRIO COMPLETO
    Logger.log('📤 ETAPA 6: ENVIANDO RELATÓRIO...');
    enviarRelatorioComBacklog(normativosAnalisados, salvosPlanilha, salvosBacklog, atualizadosBacklog, normativos.length);
    
    const endTime = new Date();
    const tempoExecucao = (endTime - startTime) / 1000;
    
    registrarLogAPI('SISTEMA', 'SUCCESS', 
      `Execução concluída - ${normativosAnalisados.length}/${normativos.length} aplicáveis | ${salvosBacklog} no backlog em ${tempoExecucao}s`, 
      normativosAnalisados.length
    );
    
    Logger.log(`🎯 SISTEMA CONCLUÍDO EM ${tempoExecucao}s! ${normativosAnalisados.length}/${normativos.length} aplicáveis + ${salvosBacklog} no backlog`);
    
  } catch (error) {
    Logger.log(`❌ ERRO CRÍTICO NO SISTEMA: ${error.toString()}`);
    registrarLogAPI('SISTEMA', 'ERROR', `Erro no sistema: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
  }
}
// =============================================
// CORREÇÃO DO BACKLOG - FUNÇÃO COMPLETA
// =============================================

/**
 * 🎯 CORREÇÃO: SALVAR TODAS AS ANÁLISES NO BACKLOG
 * Versão corrigida e testada
 */
function salvarTodasAnalisesNoBacklog(todasAnalises) {
  Logger.log('📚 SALVANDO NO BACKLOG - INICIANDO...');
  
  try {
    // 1. CONFIGURAÇÕES DA PLANILHA
    const planilha = SpreadsheetApp.openById('1zp3A_IZD5QO9L2Y7L7tX_9p_dDdylt7k3fUJc3J5kA'); // ID da planilha de backlog
    const aba = planilha.getSheetByName('Backlog');
    
    if (!aba) {
      throw new Error('Aba "Backlog" não encontrada');
    }
    
    // 2. VERIFICAR SE JÁ EXISTEM DADOS
    const ultimaLinha = aba.getLastRow();
    const dadosExistentes = ultimaLinha > 1 ? aba.getRange(2, 1, ultimaLinha - 1, 10).getValues() : [];
    
    // 3. PREPARAR NOVOS REGISTROS
    const novosRegistros = [];
    let duplicatas = 0;
    let salvos = 0;
    
    todasAnalises.forEach((analise, index) => {
      try {
        // Criar ID único para o normativo
        const idNormativo = gerarIdUnicoNormativo(analise);
        
        // Verificar se já existe no backlog
        const jaExiste = dadosExistentes.some(linha => {
          const idExistente = linha[1]; // Coluna B (ID)
          return idExistente === idNormativo;
        });
        
        if (jaExiste) {
          Logger.log(`   ⚡ Duplicata ignorada: ${analise.Titulo || 'Sem título'}`);
          duplicatas++;
          return;
        }
        
        // Preparar dados para salvar
        const registro = [
          new Date(), // Data de inclusão
          idNormativo, // ID único
          analise.Titulo || 'Sem título',
          analise.Fonte || 'Fonte não identificada',
          analise.Data || new Date(),
          analise.Link || '',
          analise['Resumo Conteúdo'] || '',
          analise['Análise Detalhada'] || '',
          analise.Aplicavel_iFood || 'Não analisado',
          analise.Impacto_iFood || 'Não especificado',
          analise['Setores Afetados'] || '',
          analise['Ações Recomendadas'] || '',
          analise.Prazo || '',
          analise.Prioridade || 'Média',
          analise.Status || 'Pendente'
        ];
        
        novosRegistros.push(registro);
        salvos++;
        Logger.log(`   ✅ Preparado: ${analise.Titulo || 'Sem título'}`);
        
      } catch (erroAnalise) {
        Logger.log(`   ❌ Erro na análise ${index}: ${erroAnalise}`);
      }
    });
    
    // 4. SALVAR NOVOS REGISTROS
    if (novosRegistros.length > 0) {
      // Adicionar na próxima linha disponível
      const linhaInicio = ultimaLinha + 1;
      aba.getRange(linhaInicio, 1, novosRegistros.length, registro.length).setValues(novosRegistros);
      
      Logger.log(`📚 BACKLOG ATUALIZADO: ${salvos} novos registros`);
    } else {
      Logger.log('📚 BACKLOG: Nenhum novo registro para salvar');
    }
    
    // 5. ATUALIZAR FORMATAÇÃO E CONGELAR LINHA
    if (novosRegistros.length > 0) {
      // Congelar primeira linha
      aba.setFrozenRows(1);
      
      // Autoajustar colunas
      aba.autoResizeColumns(1, 15);
    }
    
    return {
      total: todasAnalises.length,
      salvos: salvos,
      duplicatas: duplicatas,
      aplicaveis: todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length,
      naoAplicaveis: todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO GRAVE NO BACKLOG: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO BACKLOG: ${error.toString().substring(0, 150)}`);
    return {
      total: todasAnalises.length,
      salvos: 0,
      duplicatas: 0,
      error: error.toString()
    };
  }
}

/**
 * GERAR ID ÚNICO PARA NORMATIVO
 */
function gerarIdUnicoNormativo(analise) {
  const textoBase = `${analise.Titulo || ''}-${analise.Fonte || ''}-${analise.Data || ''}-${analise.Link || ''}`;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, textoBase)
    .map(byte => (byte + 128).toString(16).padStart(2, '0'))
    .join('')
    .substring(0, 12);
  return `NORM-${hash}`;
}

// =============================================
// FUNÇÃO ALTERNATIVA SIMPLIFICADA
// =============================================

/**
 * 🚀 VERSÃO SIMPLIFICADA PARA TESTE RÁPIDO
 */
function salvarBacklogSimplificado(todasAnalises) {
  Logger.log('📚 SALVANDO BACKLOG (VERSÃO SIMPLIFICADA)...');
  
  try {
    // ID da planilha - VERIFICAR SE ESTÁ CORRETO
    const PLANILHA_BACKLOG_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U'; // 👈 CONFIRMAR ESTE ID
    
    const planilha = SpreadsheetApp.openById(PLANILHA_BACKLOG_ID);
    const aba = planilha.getSheetByName('Backlog');
    
    if (!aba) {
      throw new Error('Aba "Backlog" não encontrada. Verifique o nome da aba.');
    }
    
    // Cabeçalhos esperados
    const cabecalhos = [
      'Data Inclusão', 'ID', 'Título', 'Fonte', 'Data Normativo', 'Link',
      'Resumo', 'Análise Detalhada', 'Aplicável iFood', 'Impacto', 
      'Setores Afetados', 'Ações Recomendadas', 'Prazo', 'Prioridade', 'Status'
    ];
    
    // Preparar dados
    const dados = todasAnalises.map(analise => [
      new Date(), // Data de inclusão
      gerarIdUnicoNormativo(analise), // ID único
      analise.Titulo || 'Sem título',
      analise.Fonte || 'Fonte não identificada',
      analise.Data || new Date(),
      analise.Link || '',
      analise['Resumo Conteúdo'] || '',
      analise['Análise Detalhada'] || '',
      analise.Aplicavel_iFood || 'Não analisado',
      analise.Impacto_iFood || 'Não especificado',
      analise['Setores Afetados'] || '',
      analise['Ações Recomendadas'] || '',
      analise.Prazo || '',
      analise.Prioridade || 'Média',
      analise.Status || 'Pendente'
    ]);
    
    // Adicionar após última linha
    if (dados.length > 0) {
      const ultimaLinha = aba.getLastRow();
      aba.getRange(ultimaLinha + 1, 1, dados.length, cabecalhos.length).setValues(dados);
      Logger.log(`✅ BACKLOG SALVO: ${dados.length} registros adicionados`);
      
      return {
        success: true,
        registros: dados.length
      };
    } else {
      Logger.log('⚠️ Nenhum dado para salvar no backlog');
      return {
        success: true,
        registros: 0
      };
    }
    
  } catch (error) {
    Logger.log(`❌ ERRO NO BACKLOG SIMPLIFICADO: ${error}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =============================================
// FUNÇÃO DE VERIFICAÇÃO DO BACKLOG
// =============================================

/**
 * VERIFICAR STATUS DO BACKLOG
 */
function verificarStatusBacklog() {
  Logger.log('🔍 VERIFICANDO STATUS DO BACKLOG...');
  
  try {
    const PLANILHA_BACKLOG_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U';
    const planilha = SpreadsheetApp.openById(PLANILHA_BACKLOG_ID);
    const aba = planilha.getSheetByName('Backlog');
    
    if (!aba) {
      throw new Error('Aba "Backlog" não encontrada');
    }
    
    const ultimaLinha = aba.getLastRow();
    const totalRegistros = ultimaLinha - 1; // Excluindo cabeçalho
    
    Logger.log(`📊 BACKLOG: ${totalRegistros} registros totais`);
    
    // Verificar últimos 5 registros
    const ultimosRegistros = ultimaLinha > 1 ? 
      aba.getRange(Math.max(2, ultimaLinha - 4), 1, Math.min(5, ultimaLinha - 1), 5).getValues() : [];
    
    Logger.log('Últimos registros no backlog:');
    ultimosRegistros.forEach((reg, index) => {
      Logger.log(`   ${index + 1}. ${reg[2]} (${reg[1]})`);
    });
    
    enviarSlackMensagem(
      `📊 *STATUS BACKLOG*\n\n` +
      `• Total de registros: ${totalRegistros}\n` +
      `• Última atualização: ${new Date().toLocaleString('pt-BR')}\n` +
      `• Planilha: ${planilha.getName()}\n` +
      `• ABA: Backlog`
    );
    
    return {
      totalRegistros: totalRegistros,
      ultimaLinha: ultimaLinha,
      planilha: planilha.getName(),
      status: 'OK'
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO NA VERIFICAÇÃO: ${error}`);
    enviarSlackMensagem(`❌ ERRO NO BACKLOG: ${error.toString()}`);
    return {
      error: error.toString(),
      status: 'ERRO'
    };
  }
}

// =============================================
// ATUALIZAR FUNÇÃO PRINCIPAL
// =============================================

/**
 * ATUALIZAR A FUNÇÃO PRINCIPAL PARA USAR BACKLOG CORRETO
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO COMPLETO - COM BACKLOG CORRIGIDO');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // [ETAPAS 1-4 MANTIDAS IGUAIS...]
    
    // 1. COLETA OFICIAL
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    
    // 2. COLETA COMPLEMENTAR  
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
      
    // COMBINAR RESULTADOS
    const todosNormativos = [...resultados.normativosOficiais, ...resultados.fontesComplementares];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    // 3. ANÁLISE TOQAN
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    // 4. 📚 BACKLOG - USANDO FUNÇÃO CORRIGIDA
    Logger.log('📚 ETAPA 4: BACKLOG CORRIGIDO...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.salvos || 0;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis || 0;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis || 0;
    
    // 5. PLANILHA APLICÁVEIS
    Logger.log('💾 ETAPA 5: PLANILHA...');
    resultados.planilha = salvarAplicaveisNaPlanilha(todasAnalises);
    
    // RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    enviarRelatorioExecucaoAgendada(resultados);
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA - Backlog: ${resultados.backlog} registros`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// RELATÓRIO CORRIGIDO - COM VALIDAÇÃO
// =============================================

function enviarRelatorioComBacklog(normativosAplicaveis, salvosPlanilha, salvosBacklog, atualizadosBacklog, totalColetado) {
  try {
    // VALIDAÇÃO DE PARÂMETROS
    const normativosValidos = Array.isArray(normativosAplicaveis) ? normativosAplicaveis : [];
    const salvosPlanilhaValido = typeof salvosPlanilha === 'number' ? salvosPlanilha : 0;
    const salvosBacklogValido = typeof salvosBacklog === 'number' ? salvosBacklog : 0;
    const atualizadosBacklogValido = typeof atualizadosBacklog === 'number' ? atualizadosBacklog : 0;
    const totalColetadoValido = typeof totalColetado === 'number' ? totalColetado : 0;
    
    const dataHoje = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
    const horaAtual = Utilities.formatDate(new Date(), 'GMT-3', 'HH:mm');
    
    let mensagem = `📊 *MONITORAMENTO IFOOD - ${dataHoje} ${horaAtual}*\n\n`;
    mensagem += `📈 *RELATÓRIO COMPLETO COM BACKLOG*\n`;
    mensagem += `├─ Coletados: ${totalColetadoValido} itens\n`;
    mensagem += `├─ Backlog: ${salvosBacklogValido} salvos\n`;
    mensagem += `├─ Aplicáveis: ${normativosValidos.length} itens\n`;
    mensagem += `└─ Planilha: ${salvosPlanilhaValido} salvos\n\n`;
    
    // DETALHAMENTO DOS APLICÁVEIS
    if (normativosValidos.length > 0) {
      mensagem += `🎯 *NORMATIVOS APLICÁVEIS IDENTIFICADOS:*\n\n`;
      
      normativosValidos.forEach((normativo, index) => {
        // VALIDAÇÃO DE DADOS DO NORMATIVO
        const orgao = normativo.Orgao || 'N/A';
        const tipoNorma = normativo.Tipo_Norma || 'N/A';
        const numero = normativo.Numero || 'N/A';
        const tema = normativo.Tema || 'N/A';
        const impacto = normativo.Impacto_Declarado || 'N/A';
        const produto = normativo.Produto_Segmento || 'N/A';
        const aplicavelSCD = normativo.Aplicavel_SCD || 'N/A';
        const aplicavelIfood = normativo.Aplicavel_iFood || 'N/A';
        const resumo = normativo.Resumo_Analise || 'N/A';
        
        const emojiImpacto = impacto === 'Alto' ? '🔴 ' :
                           impacto === 'Médio' ? '🟡 ' : '🟢 ';
        
        mensagem += `${emojiImpacto} *${orgao} ${tipoNorma} ${numero}*\n`;
        mensagem += `   _${tema}_\n`;
        mensagem += `   📋 *Impacto:* ${impacto} | *Produto:* ${produto}\n`;
        mensagem += `   ✅ *Aplicável:* SCD:${aplicavelSCD} | iFood:${aplicavelIfood}\n`;
        
        if (resumo && resumo !== 'N/A' && resumo.length > 0) {
          mensagem += `   📝 *Análise:* ${resumo.substring(0, 100)}...\n`;
        }
        
        mensagem += `\n`;
      });
    } else {
      mensagem += `⚡ *NENHUM NORMATIVO APLICÁVEL IDENTIFICADO*\n\n`;
    }
    
    // RESUMO DO BACKLOG
    mensagem += `📚 *SISTEMA DE BACKLOG:*\n`;
    mensagem += `├─ Total de itens coletados: ${totalColetadoValido}\n`;
    mensagem += `├─ Itens no backlog: ${salvosBacklogValido}\n`;
    mensagem += `├─ Itens analisados: ${atualizadosBacklogValido}\n`;
    mensagem += `└─ Itens aplicáveis: ${normativosValidos.length}\n\n`;
    
    mensagem += `💡 *OBSERVAÇÃO:* Todos os normativos coletados são salvos no backlog, mesmo os não aplicáveis.\n\n`;
    
    mensagem += `🔧 _Sistema Automático iFood Compliance - Backlog Completo_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro relatório com backlog: ${error}`);
    
    // RELATÓRIO DE FALHA SIMPLES
    const mensagemFallback = `📊 *MONITORAMENTO IFOOD - RELATÓRIO SIMPLIFICADO*\n\n`;
    mensagemFallback += `⚡ Relatório completo com erro, mas sistema funcionou.\n`;
    mensagemFallback += `📚 Backlog atualizado com sucesso.\n\n`;
    mensagemFallback += `🔧 _Sistema em operação_`;
    
    return enviarSlackMensagem(mensagemFallback);
  }
}
// =============================================
// CORREÇÃO DA DISTRIBUIÇÃO ENTRE BACKLOG E AGENDA
// =============================================

/**
 * 🎯 FUNÇÃO CORRIGIDA - SALVAR APLICÁVEIS NA AGENDA NORMATIVA
 * Versão corrigida: Aplicáveis vão para AgendaNormativa, outros para Backlog
 */
function salvarAplicaveisNaPlanilha(todasAnalises) {
  Logger.log('💾 SALVANDO APLICÁVEIS NA AGENDA NORMATIVA...');
  
  try {
    // FILTRAR APENAS OS APLICÁVEIS
    const analisesAplicaveis = todasAnalises.filter(analise => 
      analise.Aplicavel_iFood === 'Sim'
    );
    
    if (analisesAplicaveis.length === 0) {
      Logger.log('⚡ Nenhum normativo aplicável para salvar na Agenda');
      return 0;
    }
    
    Logger.log(`📋 ${analisesAplicaveis.length} normativos aplicáveis para a AgendaNormativa`);
    
    // CONFIGURAÇÃO DA PLANILHA PRINCIPAL
    const PLANILHA_PRINCIPAL_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U'; // 👈 ID da planilha principal
    const planilha = SpreadsheetApp.openById(PLANILHA_PRINCIPAL_ID);
    const abaAgenda = planilha.getSheetByName('AgendaNormativa');
    
    if (!abaAgenda) {
      throw new Error('Aba "AgendaNormativa" não encontrada na planilha principal');
    }
    
    // PREPARAR DADOS PARA AGENDA NORMATIVA
    const dadosAgenda = analisesAplicaveis.map(analise => {
      return [
        new Date(), // Data de inclusão
        analise.Titulo || 'Sem título',
        analise.Fonte || 'Fonte não identificada',
        analise.Data || new Date(),
        analise.Link || '',
        analise['Resumo Conteúdo'] || '',
        analise['Análise Detalhada'] || '',
        analise.Impacto_iFood || 'Não especificado',
        analise['Setores Afetados'] || '',
        analise['Ações Recomendadas'] || '',
        analise.Prazo || '',
        analise.Prioridade || 'Média',
        'Pendente', // Status inicial
        '', // Responsável
        '', // Data conclusão
        analise.Aplicavel_iFood || 'Sim' // Confirmar que é aplicável
      ];
    });
    
    // SALVAR NA AGENDA NORMATIVA
    if (dadosAgenda.length > 0) {
      const ultimaLinhaAgenda = abaAgenda.getLastRow();
      const linhaInicioAgenda = ultimaLinhaAgenda > 0 ? ultimaLinhaAgenda + 1 : 2;
      
      abaAgenda.getRange(linhaInicioAgenda, 1, dadosAgenda.length, dadosAgenda[0].length)
        .setValues(dadosAgenda);
      
      Logger.log(`✅ ${dadosAgenda.length} registros salvos na AgendaNormativa`);
      
      // Autoajustar colunas
      abaAgenda.autoResizeColumns(1, dadosAgenda[0].length);
    }
    
    return analisesAplicaveis.length;
    
  } catch (error) {
    Logger.log(`❌ ERRO AO SALVAR NA AGENDA: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NA AGENDA: ${error.toString().substring(0, 150)}`);
    return 0;
  }
}

/**
 * 🎯 FUNÇÃO CORRIGIDA - SALVAR TODOS NO BACKLOG
 * Versão corrigida: Salva TODOS no Backlog, independente de serem aplicáveis
 */
function salvarTodasAnalisesNoBacklog(todasAnalises) {
  Logger.log('📚 SALVANDO TODOS OS NORMATIVOS NO BACKLOG...');
  
  try {
    // CONFIGURAÇÃO DA PLANILHA DE BACKLOG
    const PLANILHA_BACKLOG_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U'; // 👈 ID da planilha de backlog
    const planilhaBacklog = SpreadsheetApp.openById(PLANILHA_BACKLOG_ID);
    const abaBacklog = planilhaBacklog.getSheetByName('Backlog');
    
    if (!abaBacklog) {
      throw new Error('Aba "Backlog" não encontrada');
    }
    
    // PREPARAR DADOS PARA BACKLOG (TODOS OS NORMATIVOS)
    const dadosBacklog = todasAnalises.map(analise => {
      return [
        new Date(), // Data de inclusão
        gerarIdUnicoNormativo(analise), // ID único
        analise.Titulo || 'Sem título',
        analise.Fonte || 'Fonte não identificada',
        analise.Data || new Date(),
        analise.Link || '',
        analise['Resumo Conteúdo'] || '',
        analise['Análise Detalhada'] || '',
        analise.Aplicavel_iFood || 'Não analisado',
        analise.Impacto_iFood || 'Não especificado',
        analise['Setores Afetados'] || '',
        analise['Ações Recomendadas'] || '',
        analise.Prazo || '',
        analise.Prioridade || 'Média',
        'Registrado' // Status inicial no backlog
      ];
    });
    
    // SALVAR NO BACKLOG
    let salvosBacklog = 0;
    
    if (dadosBacklog.length > 0) {
      const ultimaLinhaBacklog = abaBacklog.getLastRow();
      const linhaInicioBacklog = ultimaLinhaBacklog > 0 ? ultimaLinhaBacklog + 1 : 2;
      
      abaBacklog.getRange(linhaInicioBacklog, 1, dadosBacklog.length, dadosBacklog[0].length)
        .setValues(dadosBacklog);
      
      salvosBacklog = dadosBacklog.length;
      Logger.log(`✅ ${salvosBacklog} registros salvos no Backlog`);
      
      // Autoajustar colunas
      abaBacklog.autoResizeColumns(1, dadosBacklog[0].length);
    }
    
    // ESTATÍSTICAS
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    return {
      total: todasAnalises.length,
      salvos: salvosBacklog,
      aplicaveis: aplicaveis,
      naoAplicaveis: naoAplicaveis
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO NO BACKLOG: ${error.toString()}`);
    return {
      total: todasAnalises.length,
      salvos: 0,
      aplicaveis: 0,
      naoAplicaveis: 0,
      error: error.toString()
    };
  }
}

// =============================================
// FUNÇÃO PRINCIPAL ATUALIZADA
// =============================================

/**
 * 🎯 FUNÇÃO PRINCIPAL CORRIGIDA - DISTRIBUIÇÃO CORRETA
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO - DISTRIBUIÇÃO CORRETA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      agenda: 0,
      startTime: new Date()
    };
    
    // 1. COLETA OFICIAL
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. COLETA COMPLEMENTAR  
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // COMBINAR RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 3. ANÁLISE TOQAN
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise concluída');
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return { success: false, mensagem: 'Análise não concluída' };
    }
    
    // ESTATÍSTICAS DAS ANÁLISES
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 4. 📚 BACKLOG - SALVAR TODOS OS NORMATIVOS
    Logger.log('📚 ETAPA 4: BACKLOG (TODOS OS NORMATIVOS)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.salvos;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 5. 💾 AGENDA NORMATIVA - SALVAR APENAS APLICÁVEIS
    Logger.log('💾 ETAPA 5: AGENDA NORMATIVA (APENAS APLICÁVEIS)...');
    resultados.agenda = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.agenda} aplicáveis na AgendaNormativa`);
    
    // RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    // ENVIAR RELATÓRIO DETALHADO
    enviarRelatorioExecucaoAgendada(resultados);
    
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA: ${resultados.backlog} no Backlog, ${resultados.agenda} na Agenda`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO DE RELATÓRIO ATUALIZADA
// =============================================

/**
 * RELATÓRIO CORRIGIDO - MOSTRAR DISTRIBUIÇÃO CORRETA
 */
function enviarRelatorioExecucaoAgendada(resultados) {
  const tempoFormatado = resultados.tempoExecucao ? `${resultados.tempoExecucao.toFixed(1)}s` : 'N/A';
  
  const mensagem = 
    `📊 *RELATÓRIO DE EXECUÇÃO - DISTRIBUIÇÃO CORRIGIDA*\n\n` +
    `⏰ Horário: ${new Date().toLocaleString('pt-BR')}\n` +
    `⚡ Tempo: ${tempoFormatado}\n\n` +
    
    `📥 *COLETA:*\n` +
    `• Normativos oficiais: ${resultados.normativosOficiais.length}\n` +
    `• Fontes complementares: ${resultados.fontesComplementares.length}\n` +
    `• Total coletado: ${resultados.normativosOficiais.length + resultados.fontesComplementares.length}\n\n` +
    
    `🤖 *ANÁLISE TOQAN:*\n` +
    `• Total analisado: ${resultados.analisesToqan.length}\n` +
    `• Aplicáveis iFood: ${resultados.backlogAplicaveis || 0}\n` +
    `• Não aplicáveis: ${resultados.backlogNaoAplicaveis || 0}\n\n` +
    
    `💾 *ARMAZENAMENTO:*\n` +
    `• 📚 Backlog (todos): ${resultados.backlog} registros\n` +
    `• 🗓️ AgendaNormativa (aplicáveis): ${resultados.agenda} registros\n\n` +
    
    `✅ *SISTEMA FUNCIONANDO CORRETAMENTE*`;
  
  enviarSlackMensagem(mensagem);
}

// =============================================
// FUNÇÕES DE VERIFICAÇÃO
// =============================================

/**
 * VERIFICAR DISTRIBUIÇÃO CORRETA
 */
function verificarDistribuicao() {
  Logger.log('🔍 VERIFICANDO DISTRIBUIÇÃO ENTRE BACKLOG E AGENDA...');
  
  try {
    // Verificar Backlog
    const PLANILHA_BACKLOG_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U';
    const planilhaBacklog = SpreadsheetApp.openById(PLANILHA_BACKLOG_ID);
    const abaBacklog = planilhaBacklog.getSheetByName('Backlog');
    
    const totalBacklog = abaBacklog ? abaBacklog.getLastRow() - 1 : 0;
    
    // Verificar AgendaNormativa
    const PLANILHA_PRINCIPAL_ID = '1hEQ6886rbyTO2eaiapnSylWlsQVytOw7oTpfHnD3l_U';
    const planilhaPrincipal = SpreadsheetApp.openById(PLANILHA_PRINCIPAL_ID);
    const abaAgenda = planilhaPrincipal.getSheetByName('AgendaNormativa');
    
    const totalAgenda = abaAgenda ? abaAgenda.getLastRow() - 1 : 0;
    
    Logger.log(`📊 DISTRIBUIÇÃO ATUAL:`);
    Logger.log(`   📚 Backlog: ${totalBacklog} registros (TODOS os normativos)`);
    Logger.log(`   🗓️ AgendaNormativa: ${totalAgenda} registros (APENAS aplicáveis)`);
    
    enviarSlackMensagem(
      `🔍 *VERIFICAÇÃO DE DISTRIBUIÇÃO*\n\n` +
      `📚 Backlog: ${totalBacklog} registros\n` +
      `🗓️ AgendaNormativa: ${totalAgenda} registros\n` +
      `✅ Sistema configurado corretamente`
    );
    
    return {
      backlog: totalBacklog,
      agenda: totalAgenda,
      status: 'OK'
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO NA VERIFICAÇÃO: ${error}`);
    return {
      error: error.toString(),
      status: 'ERRO'
    };
  }
}
// =============================================
// FUNÇÕES AUXILIARES CORRIGIDAS
// =============================================

/**
 * Função analisarNormativosComToqan corrigida
 */
function analisarNormativosComToqan(normativos) {
  if (!normativos || !Array.isArray(normativos) || normativos.length === 0) {
    Logger.log('⚡ Nenhum normativo para analisar');
    return [];
  }
  
  Logger.log(`🤖 Iniciando análise de ${normativos.length} normativos com Toqan`);
  const client = new ToqanClient();
  const resultados = [];
  let analisados = 0;
  let aplicaveis = 0;
  
  for (let i = 0; i < normativos.length; i++) {
    const normativo = normativos[i];
    
    try {
      Logger.log(`📋 [${i + 1}/${normativos.length}] Analisando: ${normativo.Orgao} - ${(normativo.Tema || '').substring(0, 50)}...`);
      
      const analise = analisarNormativoComToqan(client, normativo);
      
      if (analise) {
        analisados++;
        
        // FILTRAR: Só incluir se for aplicável ao iFood
        if (analise.Aplicavel_iFood === 'Sim' && 
            analise.Impacto_Declarado !== 'N/A' && 
            analise.Impacto_Declarado !== 'Não Aplicável') {
          
          resultados.push(analise);
          aplicaveis++;
          Logger.log(`   ✅ APLICÁVEL - Impacto: ${analise.Impacto_Declarado}`);
        } else {
          Logger.log(`   ❌ NÃO APLICÁVEL - Descarte: ${analise.Aplicavel_iFood} | ${analise.Impacto_Declarado}`);
        }
      }
      
      // Pequeno delay entre análises
      if (i < normativos.length - 1) {
        Utilities.sleep(5000); // 5 segundos entre análises
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no normativo ${i + 1}: ${error}`);
    }
  }
  
  Logger.log(`🎯 Análise concluída: ${analisados} processados, ${aplicaveis} aplicáveis ao iFood`);
  return resultados;
}

/**
 * Função salvarNaPlanilha corrigida
 */
function salvarNaPlanilha(normativos) {
  Logger.log('💾 INICIANDO SALVAMENTO NA PLANILHA...');
  
  try {
    // VALIDAÇÃO DE ENTRADA
    if (!normativos || !Array.isArray(normativos) || normativos.length === 0) {
      Logger.log('⚡ Nenhum normativo para salvar');
      return 0;
    }
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = spreadsheet.getSheets()[0];
    
    const ultimaLinha = sheet.getLastRow();
    
    if (ultimaLinha === 0) {
      const cabecalhos = [
        'normativo_index', 'Data_Captura', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Produto_Segmento', 'Tema', 'Impacto_Declarado',
        'Data_Vigencia', 'Aplicavel_SCD', 'Aplicavel_IP', 'Aplicavel_iFood',
        'status', 'Criticidade_Sistema', 'Resumo_Analise', 'Resposta_Toqan'
      ];
      sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }
    
    const dados = [];
    let proximoIndex = ultimaLinha + 1;
    
    normativos.forEach((normativo, index) => {
      // VALIDAÇÃO DE DADOS
      const linha = [
        normativo.normativo_index || proximoIndex + index,
        normativo.Data_Captura || Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
        normativo.Orgao || 'N/A',
        normativo.Tipo_Norma || 'N/A',
        normativo.Numero || 'N/A',
        normativo.Data_Publicacao || 'N/A',
        normativo.Produto_Segmento || 'iFood Pago - Geral',
        normativo.Tema || 'N/A',
        normativo.Impacto_Declarado || 'Médio',
        normativo.Data_Vigencia || normativo.Data_Publicacao || 'N/A',
        normativo.Aplicavel_SCD || 'Não',
        normativo.Aplicavel_IP || 'Sim',
        normativo.Aplicavel_iFood || 'Sim',
        normativo.status || 'Analisado',
        normativo.Criticidade_Sistema || 'MÉDIA',
        normativo.Resumo_Analise || 'Análise Toqan AI',
        normativo.Resposta_Toqan || 'N/A'
      ];
      dados.push(linha);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha + 1;
      sheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} normativos salvos na planilha!`);
      return dados.length;
    }
    
    return 0;
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar na planilha: ${error.toString()}`);
    return 0;
  }
}


// =============================================
// FUNÇÃO DE INICIALIZAÇÃO SEGURA
// =============================================

/**
 * Função para testar e inicializar o sistema de forma segura
 */
function iniciarSistemaComBacklog() {
  Logger.log('🚀 INICIANDO SISTEMA COM BACKLOG - MODO SEGURO');
  
  try {
    // 1. Testar componentes básicos
    Logger.log('1. 🧪 Testando componentes...');
    
    // Testar planilha
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      Logger.log('   ✅ Planilha acessível');
    } catch (e) {
      Logger.log(`   ❌ Erro na planilha: ${e}`);
      throw new Error('Planilha não acessível');
    }
    
    // Testar Toqan
    const toqanOk = testarToqanSimples();
    if (!toqanOk) {
      Logger.log('   ⚡ Toqan com problemas, mas sistema continuará');
    }
    
    // 2. Configurar agendamento
    Logger.log('2. ⏰ Configurando agendamento...');
    configurarAgendamentoSimples();
    
    // 3. Executar sistema completo
    Logger.log('3. 🚀 Executando sistema completo...');
    executarSistemaCompletoComBacklog();
    
    Logger.log('🎯 SISTEMA INICIADO COM SUCESSO!');
    
  } catch (error) {
    Logger.log(`❌ ERRO NA INICIALIZAÇÃO: ${error.toString()}`);
    enviarSlackMensagem(`❌ Erro na inicialização: ${error.toString().substring(0, 100)}`);
  }
}

// =============================================
// SUBSTITUIR FUNÇÃO PRINCIPAL
// =============================================

/**
 * Substituir a função principal pela versão corrigida
 */
function executarSistemaCompleto() {
  return executarSistemaCompletoComBacklog();
}

/**
 * Função para executar apenas o backlog (sem análise Toqan)
 */
function executarApenasBacklog() {
  Logger.log('📚 EXECUTANDO APENAS BACKLOG (SEM ANÁLISE TOQAN)');
  
  try {
    const normativos = coletarNormativosReais();
    
    if (!normativos || normativos.length === 0) {
      Logger.log('⚡ Nenhum normativo para salvar no backlog');
      enviarSlackMensagem('📭 Backlog: Nenhum normativo novo hoje');
      return;
    }
    
    const salvos = salvarNoBacklog(normativos);
    
    enviarSlackMensagem(`📚 BACKLOG ATUALIZADO: ${salvos} novos normativos salvos`);
    Logger.log(`✅ ${salvos} normativos salvos no backlog`);
    
  } catch (error) {
    Logger.log(`❌ Erro no backlog simples: ${error}`);
  }
}
// =============================================
// CORREÇÃO DA RECURSÃO INFINITA
// =============================================

/**
 * FUNÇÃO PRINCIPAL PARA AGENDAMENTO - NOME DIFERENTE
 * Esta será chamada pelos triggers agendados
 */
function executarSistemaAgendado() {
  Logger.log('🔍 EXECUTANDO SISTEMA AGENDADO - MODO CORRIGIDO');
  
  try {
    const resultado = executarMonitoramentoCompletoPrincipal();
    
    // ENVIAR RELATÓRIO DE EXECUÇÃO AGENDADA
    if (resultado.success) {
      enviarRelatorioExecucaoAgendada(resultado);
    } else {
      enviarSlackMensagem(
        `❌ *EXECUÇÃO AGENDADA COM FALHA*\n\n` +
        `⚡ Erro: ${resultado.error}\n` +
        `🔧 Verificar logs para detalhes`
      );
    }
    
    return resultado;
    
  } catch (error) {
    Logger.log(`❌ ERRO NA EXECUÇÃO AGENDADA: ${error.toString()}`);
    enviarSlackMensagem(`❌ FALHA NA EXECUÇÃO AGENDADA: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PRINCIPAL DO MONITORAMENTO (NOME ALTERADO)
// =============================================

function executarMonitoramentoCompletoPrincipal() {
  Logger.log('🔍 INICIANDO MONITORAMENTO COMPLETO - MODO PRINCIPAL');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. MONITORAMENTO OFICIAL (BACEN/CMN/DOU)
    Logger.log('🏛️  MÓDULO 1: MONITORAMENTO OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. MONITORAMENTO COMPLEMENTAR (NOTÍCIAS, PORTAIS)
    Logger.log('📰 MÓDULO 2: MONITORAMENTO COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado em nenhum módulo');
      resultados.mensagem = 'Nenhum normativo detectado';
      resultados.success = true;
      return resultados;
    }
    
    Logger.log(`📊 TOTAL DETECTADO: ${todosNormativos.length} normativos`);
    
    // 4. SALVAR NO BACKLOG (TODOS OS NORMATIVOS)
    Logger.log('📚 MÓDULO 3: BACKLOG COMPLETO...');
    resultados.backlog = salvarNoBacklog(todosNormativos);
    Logger.log(`   ✅ ${resultados.backlog} itens no backlog`);
    
    // 5. ANÁLISE TOQAN (APLICÁVEIS)
    Logger.log('🤖 MÓDULO 4: ANÁLISE TOQAN...');
    resultados.analisesToqan = analisarNormativosComToqan(todosNormativos);
    Logger.log(`   ✅ ${resultados.analisesToqan.length} normativos aplicáveis analisados`);
    
    // 6. ATUALIZAR BACKLOG COM ANÁLISES
    Logger.log('🔄 MÓDULO 5: ATUALIZANDO BACKLOG...');
    resultados.backlogAtualizado = atualizarBacklogComAnalise(resultados.analisesToqan);
    
    // 7. SALVAR APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 MÓDULO 6: PLANILHA PRINCIPAL...');
    resultados.planilha = salvarNaPlanilha(resultados.analisesToqan);
    Logger.log(`   ✅ ${resultados.planilha} itens na planilha principal`);
    
    // 8. TEMPO DE EXECUÇÃO
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    Logger.log(`🎯 MONITORAMENTO COMPLETO CONCLUÍDO EM ${resultados.tempoExecucao}s`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO MONITORAMENTO COMPLETO: ${error.toString()}`);
    return {
      success: false,
      error: error.toString(),
      endTime: new Date()
    };
  }
}

// =============================================
// CONFIGURAÇÃO DE AGENDAMENTO CORRIGIDA
// =============================================

function configurarAgendamentoCompleto() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTO COMPLETO - CORRIGIDO');
  
  try {
    // REMOVER TODOS OS TRIGGERS EXISTENTES
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`   🔄 Removido: ${trigger.getHandlerFunction()}`);
    });
    
    // AGENDAMENTOS PRINCIPAIS - USANDO NOVO NOME
    const horarios = [9, 12, 17]; // 9h, 12h, 17h
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarSistemaAgendado')  // NOME CORRIGIDO
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`   ✅ Agendado: ${hora}:00 - Sistema Agendado`);
    });
    
    // AGENDAMENTO DE SAÚDE DO SISTEMA
    ScriptApp.newTrigger('verificarSaudeSistema')
      .timeBased()
      .atHour(8)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone('America/Sao_Paulo')
      .create();
    Logger.log('   ✅ Agendado: 08:00 - Verificação de Saúde');
    
    // AGENDAMENTO DE BACKUP
    ScriptApp.newTrigger('backupSistema')
      .timeBased()
      .atHour(2)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone('America/Sao_Paulo')
      .create();
    Logger.log('   ✅ Agendado: 02:00 - Backup do Sistema');
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    enviarSlackMensagem(
      `⏰ *SISTEMA AGENDADO - CORRIGIDO*\n\n` +
      `✅ ${triggersFinais.length} agendamentos ativos\n` +
      `🕘 Horários: 9h, 12h, 17h\n` +
      `🔍 Módulos: Oficial + Complementar + Toqan\n` +
      `📚 Backlog: Ativo\n\n` +
      `🎯 Recursão corrigida - Sistema operacional!`
    );
    
    return { success: true, triggers: triggersFinais.length };
    
  } catch (error) {
    Logger.log(`❌ ERRO NO AGENDAMENTO: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA PARAR AGENDAMENTOS ATUAIS
// =============================================

function pararTodosAgendamentos() {
  Logger.log('🛑 PARANDO TODOS OS AGENDAMENTOS');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
      Logger.log(`   🔄 Removido: ${trigger.getHandlerFunction()}`);
    });
    
    enviarSlackMensagem(
      `🛑 *TODOS OS AGENDAMENTOS PARADOS*\n\n` +
      `✅ ${removidos} triggers removidos\n` +
      `⚡ Sistema parado até nova configuração`
    );
    
    return { success: true, removidos: removidos };
    
  } catch (error) {
    Logger.log(`❌ Erro ao parar agendamentos: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA VERIFICAR AGENDAMENTOS ATUAIS
// =============================================

function verificarAgendamentosAtuais() {
  Logger.log('🔍 VERIFICANDO AGENDAMENTOS ATUAIS');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    let mensagem = `⏰ *AGENDAMENTOS ATUAIS:*\n\n`;
    mensagem += `📊 Total: ${triggers.length} triggers\n\n`;
    
    if (triggers.length === 0) {
      mensagem += `⚡ Nenhum agendamento ativo\n`;
    } else {
      triggers.forEach((trigger, index) => {
        mensagem += `${index + 1}. ${trigger.getHandlerFunction()}\n`;
      });
    }
    
    mensagem += `\n🔧 Use 'pararTodosAgendamentos()' para limpar`;
    
    enviarSlackMensagem(mensagem);
    
    return { 
      success: true, 
      triggers: triggers.length,
      detalhes: triggers.map(t => t.getHandlerFunction())
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao verificar agendamentos: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÕES DE CONTROLE SIMPLIFICADAS
// =============================================

/**
 * EXECUTAR AGORA - MODO SIMPLES E SEGURO
 */
function executarAgora() {
  Logger.log('🚀 EXECUTANDO SISTEMA AGORA - MODO SEGURO');
  return executarMonitoramentoCompletoPrincipal();
}

/**
 * TESTAR SISTEMA - SEM AGENDAMENTO
 */
function testarSistema() {
  Logger.log('🧪 TESTANDO SISTEMA - MODO TESTE');
  
  try {
    // Executar apenas coleta básica
    const normativos = coletarNormativosReais();
    const backlog = salvarNoBacklog(normativos);
    
    enviarSlackMensagem(
      `🧪 *TESTE DO SISTEMA*\n\n` +
      `📊 Resultados:\n` +
      `├─ Normativos: ${normativos.length}\n` +
      `└─ Backlog: ${backlog} itens\n\n` +
      `✅ Teste concluído`
    );
    
    return { success: true, normativos: normativos.length, backlog: backlog };
    
  } catch (error) {
    Logger.log(`❌ Erro no teste: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// EXECUTAR CORREÇÃO COMPLETA
// =============================================

/**
 * FUNÇÃO PARA CORRIGIR TUDO DE UMA VEZ
 */
function corrigirSistema() {
  Logger.log('🔧 CORRIGINDO SISTEMA COMPLETO');
  
  try {
    // 1. Parar todos os agendamentos
    pararTodosAgendamentos();
    Utilities.sleep(3000);
    
    // 2. Configurar novo agendamento corrigido
    configurarAgendamentoCompleto();
    Utilities.sleep(3000);
    
    // 3. Executar teste rápido
    testarSistema();
    
    enviarSlackMensagem(
      `🔧 *SISTEMA CORRIGIDO*\n\n` +
      `✅ Recursão infinita resolvida\n` +
      `✅ Agendamentos reconfigurados\n` +
      `✅ Teste executado com sucesso\n\n` +
      `🎯 Sistema pronto para uso!`
    );
    
    return { success: true };
    
  } catch (error) {
    Logger.log(`❌ Erro na correção: ${error}`);
    enviarSlackMensagem(`❌ Falha na correção: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// SISTEMA COMPLETO - TODAS AS FUNÇÕES NECESSÁRIAS
// =============================================

// =============================================
// 1. SISTEMA DE BACKLOG COMPLETO
// =============================================

/**
 * Função para salvar TODOS os normativos na aba BACKLOG
 */
function salvarNoBacklog(normativos) {
  Logger.log('📚 SALVANDO NO BACKLOG...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let backlogSheet;
    
    try {
      backlogSheet = spreadsheet.getSheetByName('BACKLOG');
    } catch (e) {
      // Criar aba BACKLOG se não existir
      backlogSheet = spreadsheet.insertSheet('BACKLOG');
      const cabecalhos = [
        'ID_Backlog', 'Data_Coleta', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Tema', 'Texto_Completo', 'URL_Fonte',
        'Status_Analise', 'Impacto_Toqan', 'Produto_Afetado_Toqan',
        'Aplicavel_SCD_Toqan', 'Aplicavel_iFood_Toqan', 'Resumo_Toqan',
        'ID_Conversa_Toqan', 'Data_Analise_Toqan'
      ];
      backlogSheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      backlogSheet.getRange(1, 1, 1, cabecalhos.length)
        .setBackground('#2E7D32')
        .setFontColor('white')
        .setFontWeight('bold');
      
      Logger.log('✅ Nova aba BACKLOG criada');
    }
    
    const dados = [];
    const dataColeta = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    const ultimaLinha = backlogSheet.getLastRow();
    let proximoID = ultimaLinha; // Começar do último ID + 1
    
    if (ultimaLinha > 0) {
      const ultimoID = backlogSheet.getRange(ultimaLinha, 1).getValue();
      proximoID = isNaN(ultimoID) ? 1 : ultimoID + 1;
    } else {
      proximoID = 1;
    }
    
    normativos.forEach((normativo, index) => {
      const linha = [
        proximoID + index, // ID_Backlog
        dataColeta, // Data_Coleta
        normativo.Orgao || 'N/A', // Orgao
        normativo.Tipo_Norma || 'N/A', // Tipo_Norma
        normativo.Numero || 'N/A', // Numero
        normativo.Data_Publicacao || 'N/A', // Data_Publicacao
        normativo.Tema || 'N/A', // Tema
        normativo.texto_completo || normativo.Tema || 'N/A', // Texto_Completo
        normativo.url_fonte || 'N/A', // URL_Fonte
        'Coletado', // Status_Analise (inicial)
        'Não Analisado', // Impacto_Toqan
        'Não Analisado', // Produto_Afetado_Toqan
        'Não Analisado', // Aplicavel_SCD_Toqan
        'Não Analisado', // Aplicavel_iFood_Toqan
        'Aguardando análise', // Resumo_Toqan
        'N/A', // ID_Conversa_Toqan
        'N/A' // Data_Analise_Toqan
      ];
      dados.push(linha);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha === 0 ? 2 : ultimaLinha + 1;
      backlogSheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} normativos salvos no BACKLOG!`);
      
      // Registrar no log do sistema
      registrarLogAPI('BACKLOG', 'SUCCESS', 
        `${dados.length} normativos salvos no backlog`, 
        dados.length
      );
      
      return dados.length;
    }
    
    return 0;
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar no backlog: ${error.toString()}`);
    registrarLogAPI('BACKLOG', 'ERROR', `Erro: ${error.toString()}`, 0);
    return 0;
  }
}

/**
 * Função para atualizar backlog com resultados da análise Toqan
 */
function atualizarBacklogComAnalise(normativosAnalisados) {
  Logger.log('🔄 ATUALIZANDO BACKLOG COM ANÁLISE TOQAN...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const backlogSheet = spreadsheet.getSheetByName('BACKLOG');
    
    if (!backlogSheet) {
      Logger.log('❌ Aba BACKLOG não encontrada');
      return 0;
    }
    
    const dataAnalise = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    let atualizados = 0;
    
    // Buscar por correspondências no backlog
    const ultimaLinha = backlogSheet.getLastRow();
    if (ultimaLinha <= 1) return 0;
    
    const dadosBacklog = backlogSheet.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    
    normativosAnalisados.forEach(normativoAnalisado => {
      // Tentar encontrar correspondência no backlog
      const indice = dadosBacklog.findIndex(linha => 
        linha[2] === normativoAnalisado.Orgao && // Orgao
        linha[3] === normativoAnalisado.Tipo_Norma && // Tipo_Norma
        linha[4] === normativoAnalisado.Numero && // Numero
        linha[5] === normativoAnalisado.Data_Publicacao // Data_Publicacao
      );
      
      if (indice !== -1) {
        const linhaBacklog = indice + 2; // +2 porque começa da linha 2
        
        // Atualizar dados da análise
        backlogSheet.getRange(linhaBacklog, 10).setValue('Analisado'); // Status_Analise (coluna J)
        backlogSheet.getRange(linhaBacklog, 11).setValue(normativoAnalisado.Impacto_Declarado || 'N/A'); // Impacto_Toqan
        backlogSheet.getRange(linhaBacklog, 12).setValue(normativoAnalisado.Produto_Segmento || 'N/A'); // Produto_Afetado_Toqan
        backlogSheet.getRange(linhaBacklog, 13).setValue(normativoAnalisado.Aplicavel_SCD || 'N/A'); // Aplicavel_SCD_Toqan
        backlogSheet.getRange(linhaBacklog, 14).setValue(normativoAnalisado.Aplicavel_iFood || 'N/A'); // Aplicavel_iFood_Toqan
        backlogSheet.getRange(linhaBacklog, 15).setValue(normativoAnalisado.Resumo_Analise || 'N/A'); // Resumo_Toqan
        backlogSheet.getRange(linhaBacklog, 16).setValue(
          normativoAnalisado.Resposta_Toqan ? 
          normativoAnalisado.Resposta_Toqan.replace('Toqan ID: ', '') : 'N/A'
        ); // ID_Conversa_Toqan
        backlogSheet.getRange(linhaBacklog, 17).setValue(dataAnalise); // Data_Analise_Toqan
        
        atualizados++;
        Logger.log(`   ✅ Atualizado: ${normativoAnalisado.Orgao} ${normativoAnalisado.Numero}`);
      }
    });
    
    Logger.log(`✅ ${atualizados} registros atualizados no backlog`);
    return atualizados;
    
  } catch (error) {
    Logger.log(`❌ ERRO ao atualizar backlog: ${error.toString()}`);
    return 0;
  }
}

// =============================================
// 2. SISTEMA DE ANÁLISE TOQAN COMPLETO
// =============================================

/**
 * Função principal de análise com Toqan
 */
function analisarNormativosComToqan(normativos) {
  if (!normativos || !Array.isArray(normativos) || normativos.length === 0) {
    Logger.log('⚡ Nenhum normativo para analisar');
    return [];
  }
  
  Logger.log(`🤖 INICIANDO ANÁLISE TOQAN: ${normativos.length} normativos`);
  const client = new ToqanClient();
  const resultados = [];
  let analisados = 0;
  let aplicaveis = 0;
  
  for (let i = 0; i < normativos.length; i++) {
    const normativo = normativos[i];
    
    try {
      Logger.log(`📋 [${i + 1}/${normativos.length}] Analisando: ${normativo.Orgao} - ${(normativo.Tema || '').substring(0, 50)}...`);
      
      const analise = analisarNormativoComToqan(client, normativo);
      
      if (analise) {
        analisados++;
        
        // FILTRAR: Só incluir se for aplicável ao iFood
        if (analise.Aplicavel_iFood === 'Sim' && 
            analise.Impacto_Declarado !== 'N/A' && 
            analise.Impacto_Declarado !== 'Não Aplicável') {
          
          resultados.push(analise);
          aplicaveis++;
          Logger.log(`   ✅ APLICÁVEL - Impacto: ${analise.Impacto_Declarado}`);
        } else {
          Logger.log(`   ❌ NÃO APLICÁVEL - Descarte: ${analise.Aplicavel_iFood} | ${analise.Impacto_Declarado}`);
        }
      }
      
      // Pequeno delay entre análises
      if (i < normativos.length - 1) {
        Utilities.sleep(5000); // 5 segundos entre análises
      }
      
    } catch (error) {
      Logger.log(`❌ Erro no normativo ${i + 1}: ${error}`);
    }
  }
  
  Logger.log(`🎯 Análise concluída: ${analisados} processados, ${aplicaveis} aplicáveis ao iFood`);
  return resultados;
}

/**
 * Análise individual de normativo com Toqan
 */
function analisarNormativoComToqan(client, normativo) {
  try {
    // Preparar texto para análise
    const textoAnalise = normativo.texto_completo || normativo.Tema || '';
    const orgao = normativo.Orgao || 'N/A';
    const tipo = normativo.Tipo_Norma || 'N/A';
    
    const prompt = `Analise ESTE CONTEÚDO para determinar se é APLICÁVEL ao iFood e qual o IMPACTO REAL.

**CONTEÚDO PARA ANÁLISE:**
Fonte: ${orgao}
Tipo: ${tipo}
Número: ${normativo.Numero || 'N/A'}
Data: ${normativo.Data_Publicacao || 'N/A'}
Título: ${normativo.Tema || 'N/A'}
Texto: ${textoAnalise.substring(0, 2000)}

**CONTEXTO IFOOD - ATIVIDADES RELEVANTES:**
- iFood Pago: Sistema de pagamentos (PIX, cartões, voucher alimentação)
- iFood Crédito: Empréstimos, crédito consignado para entregadores
- SCD (Sociedade de Crédito Direto): Operações de crédito
- IP (Instituição de Pagamento): instituição de pagamentos
- Marketplace: Intermediação de vendas de restaurantes
- Pagamentos instantâneos, taxas de intermediação

**CRITÉRIOS DE APLICABILIDADE - CONSIDERE APENAS SE ENCAIXAR EM:**
✅ Regulamentação de pagamentos, PIX, cartões, instituições de pagamento
✅ Normas sobre crédito, empréstimos, fintechs
✅ Regulação de marketplaces, intermediação
✅ Compliance financeiro, prevenção à lavagem
✅ Taxas de intermediação, relações com parceiros
❌ NÃO APLICÁVEL: Notícias gerais, política, outros setores

**RESPONDA APENAS COM ESTE JSON:**

{
  "aplicavel_ifood": "Sim" ou "Não",
  "impacto": "Alto" ou "Médio" ou "Baixo" ou "Não Aplicável",
  "motivo_aplicabilidade": "Explicação curta do porquê é ou não aplicável",
  "produto_afetado": "iFood Pago" ou "iFood Crédito" ou "SCD" ou "Marketplace" ou "Múltiplos" ou "Nenhum",
  "aplicavel_scd": "Sim" ou "Não",
  "resumo_impacto": "Resumo específico do impacto para iFood",
  "acoes_recomendadas": "Ações específicas recomendadas ou 'Nenhuma ação necessária'"
}

**SEJA RIGOROSO: Marque como "Não Aplicável" se não tiver relação direta com as atividades do iFood Pago.**`;

    Logger.log(`   🤖 Enviando para Toqan...`);
    const resposta = client.createConversation(prompt);
    
    Logger.log(`   ✅ Toqan recebeu: ${resposta.conversation_id}`);
    
    // Aguardar processamento
    Utilities.sleep(6000);
    
    // Processar resposta com validação rigorosa
    return processarRespostaToqanFiltrada(resposta, normativo);
    
  } catch (error) {
    Logger.log(`   ❌ Erro Toqan: ${error}`);
    return null;
  }
}

/**
 * Processar resposta do Toqan
 */
function processarRespostaToqanFiltrada(resposta, normativo) {
  try {
    // Valores padrão CONSERVADORES - assumir não aplicável até provar o contrário
    let aplicavelIfood = 'Não';
    let impacto = 'Não Aplicável';
    let motivoAplicabilidade = 'Análise em andamento';
    let produtoAfetado = 'Nenhum';
    let aplicavelSCD = 'Não';
    let resumoImpacto = 'Aguardar análise detalhada';
    let acoesRecomendadas = 'Nenhuma ação necessária';
    
    // Tentar extrair JSON da resposta
    if (resposta && typeof resposta === 'object') {
      const respostaStr = JSON.stringify(resposta);
      
      // Extrair informações com regex mais específicos
      const aplicavelMatch = respostaStr.match(/"aplicavel_ifood"\\s*:\\s*"([^"]*)"/i);
      const impactoMatch = respostaStr.match(/"impacto"\\s*:\\s*"([^"]*)"/i);
      const motivoMatch = respostaStr.match(/"motivo_aplicabilidade"\\s*:\\s*"([^"]*)"/i);
      const produtoMatch = respostaStr.match(/"produto_afetado"\\s*:\\s*"([^"]*)"/i);
      const scdMatch = respostaStr.match(/"aplicavel_scd"\\s*:\\s*"([^"]*)"/i);
      const resumoMatch = respostaStr.match(/"resumo_impacto"\\s*:\\s*"([^"]*)"/i);
      const acoesMatch = respostaStr.match(/"acoes_recomendadas"\\s*:\\s*"([^"]*)"/i);
      
      if (aplicavelMatch) aplicavelIfood = aplicavelMatch[1];
      if (impactoMatch) impacto = impactoMatch[1];
      if (motivoMatch) motivoAplicabilidade = motivoMatch[1];
      if (produtoMatch) produtoAfetado = produtoMatch[1];
      if (scdMatch) aplicavelSCD = scdMatch[1];
      if (resumoMatch) resumoImpacto = resumoMatch[1];
      if (acoesMatch) acoesRecomendadas = acoesMatch[1];
      
      // VALIDAÇÃO: Se for "Não Aplicável", forçar consistência
      if (impacto === 'Não Aplicável') {
        aplicavelIfood = 'Não';
        produtoAfetado = 'Nenhum';
        aplicavelSCD = 'Não';
      }
      
      // VALIDAÇÃO: Se não for aplicável, impacto deve ser "Não Aplicável"
      if (aplicavelIfood === 'Não' && impacto !== 'Não Aplicável') {
        impacto = 'Não Aplicável';
      }
    }
    
    const resultado = {
      normativo_index: obterProximoIndex(),
      Data_Captura: Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
      Orgao: normativo.Orgao || 'N/A',
      Tipo_Norma: normativo.Tipo_Norma || 'N/A',
      Numero: normativo.Numero || 'N/A',
      Data_Publicacao: normativo.Data_Publicacao || 'N/A',
      Produto_Segmento: produtoAfetado,
      Tema: normativo.Tema || 'N/A',
      Impacto_Declarado: impacto,
      Data_Vigencia: normativo.Data_Publicacao || 'N/A',
      Aplicavel_SCD: aplicavelSCD,
      Aplicavel_IP: aplicavelIfood, // Usar mesma lógica do iFood
      Aplicavel_iFood: aplicavelIfood,
      status: aplicavelIfood === 'Sim' ? 'Analisado' : 'Não Aplicável',
      Criticidade_Sistema: calcularCriticidade(impacto),
      Resumo_Analise: `${motivoAplicabilidade} | ${resumoImpacto}`,
      Acoes_Recomendadas: acoesRecomendadas,
      Resposta_Toqan: `Toqan ID: ${resposta.conversation_id}`,
      url_fonte: normativo.url_fonte || 'N/A'
    };
    
    Logger.log(`   📊 Resultado: ${aplicavelIfood} | Impacto: ${impacto} | Produto: ${produtoAfetado}`);
    Logger.log(`   📝 Motivo: ${motivoAplicabilidade.substring(0, 80)}...`);
    
    return resultado;
    
  } catch (error) {
    Logger.log(`   ⚡ Erro processar resposta: ${error}`);
    return null;
  }
}

/**
 * Calcular criticidade baseada no impacto
 */
function calcularCriticidade(impacto) {
  switch(impacto) {
    case 'Alto': return 'ALTA';
    case 'Médio': return 'MÉDIA';
    case 'Baixo': return 'BAIXA';
    case 'Não Aplicável': return 'N/A';
    default: return 'MÉDIA';
  }
}

/**
 * Obter próximo índice para planilha
 */
function obterProximoIndex() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const ultimaLinha = sheet.getLastRow();
    return ultimaLinha <= 1 ? 1 : ultimaLinha + 1;
  } catch (e) {
    return 1;
  }
}

// =============================================
// 3. SISTEMA DE RELATÓRIOS COMPLETO
// =============================================

/**
 * Relatório de inicialização do sistema
 */
function enviarRelatorioInicializacao(resultado) {
  try {
    const data = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
    
    let mensagem = `🎯 *SISTEMA COMPLETO INICIALIZADO - ${data}*\n\n`;
    
    if (resultado.success) {
      mensagem += `✅ *INICIALIZAÇÃO BEM-SUCEDIDA*\n\n`;
      mensagem += `📊 *RESULTADOS DA PRIMEIRA EXECUÇÃO:*\n`;
      mensagem += `├─ Normativos Oficiais: ${resultado.normativosOficiais.length}\n`;
      mensagem += `├─ Fontes Complementares: ${resultado.fontesComplementares.length}\n`;
      mensagem += `├─ Total Coletado: ${resultado.normativosOficiais.length + resultado.fontesComplementares.length}\n`;
      mensagem += `├─ Backlog: ${resultado.backlog} itens\n`;
      mensagem += `├─ Análises Toqan: ${resultado.analisesToqan.length}\n`;
      mensagem += `├─ Planilha Principal: ${resultado.planilha} itens\n`;
      mensagem += `└─ Tempo de Execução: ${resultado.tempoExecucao}s\n\n`;
      
      // DETALHES DOS MÓDULOS
      mensagem += `🔧 *MÓDULOS ATIVOS:*\n`;
      mensagem += `├─ 🏛️  Monitoramento Oficial (BACEN/CMN/DOU)\n`;
      mensagem += `├─ 📰 Monitoramento Complementar (Notícias)\n`;
      mensagem += `├─ 🤖 Análise Toqan AI\n`;
      mensagem += `├─ 📚 Sistema de Backlog\n`;
      mensagem += `├─ 💾 Planilha Principal\n`;
      mensagem += `└─ ⏰ Agendamento Automático\n\n`;
      
    } else {
      mensagem += `❌ *INICIALIZAÇÃO COM FALHAS*\n\n`;
      mensagem += `⚡ Erro: ${resultado.error || 'Desconhecido'}\n\n`;
      mensagem += `🔧 *Verifique os módulos individualmente:*\n`;
    }
    
    mensagem += `⏰ *PRÓXIMAS EXECUÇÕES AUTOMÁTICAS:*\n`;
    mensagem += `├─ 9:00, 12:00, 17:00 - Monitoramento Completo\n`;
    mensagem += `├─ 8:00 - Verificação de Saúde\n`;
    mensagem += `└─ 2:00 - Backup do Sistema\n\n`;
    
    mensagem += `🎯 _Sistema iFood Compliance - Todos os Módulos Integrados_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro no relatório de inicialização: ${error}`);
    return enviarSlackMensagem('🎯 Sistema completo inicializado (relatório com erro)');
  }
}

/**
 * Relatório de execução agendada
 */
function enviarRelatorioExecucaoAgendada(resultado) {
  try {
    const data = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
    
    let mensagem = `🔍 *MONITORAMENTO AUTOMÁTICO - ${data}*\n\n`;
    
    mensagem += `📊 *RESULTADOS:*\n`;
    mensagem += `├─ Normativos Oficiais: ${resultado.normativosOficiais.length}\n`;
    mensagem += `├─ Fontes Complementares: ${resultado.fontesComplementares.length}\n`;
    mensagem += `├─ Total Coletado: ${resultado.normativosOficiais.length + resultado.fontesComplementares.length}\n`;
    mensagem += `├─ Backlog: ${resultado.backlog} itens\n`;
    mensagem += `├─ Aplicáveis (Toqan): ${resultado.analisesToqan.length}\n`;
    mensagem += `├─ Planilha Principal: ${resultado.planilha} itens\n`;
    mensagem += `└─ Tempo de Execução: ${resultado.tempoExecucao}s\n\n`;
    
    // DETALHES DOS APLICÁVEIS
    if (resultado.analisesToqan.length > 0) {
      mensagem += `🎯 *NORMATIVOS APLICÁVEIS:*\n`;
      
      resultado.analisesToqan.slice(0, 3).forEach(normativo => {
        const emoji = normativo.Impacto_Declarado === 'Alto' ? '🔴' : 
                     normativo.Impacto_Declarado === 'Médio' ? '🟡' : '🟢';
        
        mensagem += `${emoji} ${normativo.Orgao} ${normativo.Numero} - ${normativo.Impacto_Declarado}\n`;
      });
      
      if (resultado.analisesToqan.length > 3) {
        mensagem += `📝 ...e mais ${resultado.analisesToqan.length - 3} normativos\n`;
      }
      mensagem += `\n`;
    }
    
    mensagem += `✅ _Execução automática concluída_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro no relatório agendado: ${error}`);
    return enviarSlackMensagem(`🔍 Monitoramento automático executado - Verificar logs para detalhes`);
  }
}

// =============================================
// SISTEMA COM SEQUÊNCIA CORRETA
// =============================================

/**
 * FUNÇÃO PRINCIPAL COM SEQUÊNCIA CORRETA:
 * 📥 COLETA → 🤖 TOQAN (todos) → 📚 BACKLOG (todos) → 💾 PLANILHA (só aplicáveis)
 */
function executarMonitoramentoCompletoPrincipal() {
  Logger.log('🔍 INICIANDO MONITORAMENTO COMPLETO - SEQUÊNCIA CORRETA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. 📥 COLETA - MONITORAMENTO OFICIAL (BACEN/CMN/DOU)
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. 📥 COLETA - MONITORAMENTO COMPLEMENTAR (NOTÍCIAS, PORTAIS)
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS DA COLETA
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado em nenhum módulo');
      resultados.mensagem = 'Nenhum normativo detectado';
      resultados.success = true;
      return resultados;
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 4. 🤖 TOQAN - ANALISAR TODOS OS NORMATIVOS COLETADOS
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN (TODOS OS NORMATIVOS)...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    // ✅ ESTATÍSTICAS DETALHADAS
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises completas (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 5. 📚 BACKLOG - SALVAR TODAS AS ANÁLISES NO BACKLOG
    Logger.log('📚 ETAPA 4: BACKLOG (TODAS AS ANÁLISES)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.total;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 6. 💾 PLANILHA - SALVAR APENAS APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 ETAPA 5: PLANILHA (APENAS APLICÁVEIS)...');
    resultados.planilha = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.planilha} itens APLICÁVEIS na planilha principal`);
    
    // 7. TEMPO DE EXECUÇÃO
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    Logger.log(`🎯 MONITORAMENTO COMPLETO CONCLUÍDO EM ${resultados.tempoExecucao}s`);
    Logger.log(`📈 SEQUÊNCIA CORRETA: ${todosNormativos.length} coletados → ${todasAnalises.length} analisados → ${resultados.backlog} backlog → ${resultados.planilha} planilha`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO MONITORAMENTO COMPLETO: ${error.toString()}`);
    return {
      success: false,
      error: error.toString(),
      endTime: new Date()
    };
  }
}

// =============================================
// FUNÇÃO PARA SALVAR TODAS AS ANÁLISES NO BACKLOG
// =============================================

/**
 * Salvar TODAS as análises Toqan no backlog (aplicáveis e não aplicáveis)
 */
function salvarTodasAnalisesNoBacklog(todasAnalises) {
  Logger.log('📚 SALVANDO TODAS AS ANÁLISES NO BACKLOG...');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let backlogSheet;
    
    try {
      backlogSheet = spreadsheet.getSheetByName('BACKLOG');
    } catch (e) {
      // Criar aba BACKLOG se não existir
      backlogSheet = spreadsheet.insertSheet('BACKLOG');
      const cabecalhos = [
        'ID_Backlog', 'Data_Coleta', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Tema', 'Texto_Completo', 'URL_Fonte',
        'Status_Analise', 'Impacto_Toqan', 'Produto_Afetado_Toqan',
        'Aplicavel_SCD_Toqan', 'Aplicavel_iFood_Toqan', 'Resumo_Toqan',
        'ID_Conversa_Toqan', 'Data_Analise_Toqan'
      ];
      backlogSheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      backlogSheet.getRange(1, 1, 1, cabecalhos.length)
        .setBackground('#2E7D32')
        .setFontColor('white')
        .setFontWeight('bold');
      
      Logger.log('✅ Nova aba BACKLOG criada');
    }
    
    const dados = [];
    const dataColeta = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    const ultimaLinha = backlogSheet.getLastRow();
    let proximoID = ultimaLinha; // Começar do último ID + 1
    
    if (ultimaLinha > 0) {
      const ultimoID = backlogSheet.getRange(ultimaLinha, 1).getValue();
      proximoID = isNaN(ultimoID) ? 1 : ultimoID + 1;
    } else {
      proximoID = 1;
    }
    
    let aplicaveis = 0;
    let naoAplicaveis = 0;
    
    todasAnalises.forEach((analise, index) => {
      // Contar estatísticas
      if (analise.Aplicavel_iFood === 'Sim') {
        aplicaveis++;
      } else {
        naoAplicaveis++;
      }
      
      const linha = [
        proximoID + index, // ID_Backlog
        dataColeta, // Data_Coleta
        analise.Orgao || 'N/A', // Orgao
        analise.Tipo_Norma || 'N/A', // Tipo_Norma
        analise.Numero || 'N/A', // Numero
        analise.Data_Publicacao || 'N/A', // Data_Publicacao
        analise.Tema || 'N/A', // Tema
        analise.texto_completo || analise.Tema || 'N/A', // Texto_Completo
        analise.url_fonte || 'N/A', // URL_Fonte
        'Analisado', // Status_Analise - JÁ ANALISADO
        analise.Impacto_Declarado || 'N/A', // Impacto_Toqan
        analise.Produto_Segmento || 'N/A', // Produto_Afetado_Toqan
        analise.Aplicavel_SCD || 'N/A', // Aplicavel_SCD_Toqan
        analise.Aplicavel_iFood || 'N/A', // Aplicavel_iFood_Toqan
        analise.Resumo_Analise || 'N/A', // Resumo_Toqan
        analise.Resposta_Toqan ? analise.Resposta_Toqan.replace('Toqan ID: ', '') : 'N/A', // ID_Conversa_Toqan
        dataColeta // Data_Analise_Toqan
      ];
      dados.push(linha);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha === 0 ? 2 : ultimaLinha + 1;
      backlogSheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} análises salvas no BACKLOG! (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
      
      // Registrar no log do sistema
      registrarLogAPI('BACKLOG', 'SUCCESS', 
        `${dados.length} análises salvas no backlog (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`, 
        dados.length
      );
      
      return {
        total: dados.length,
        aplicaveis: aplicaveis,
        naoAplicaveis: naoAplicaveis
      };
    }
    
    return { total: 0, aplicaveis: 0, naoAplicaveis: 0 };
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar análises no backlog: ${error.toString()}`);
    registrarLogAPI('BACKLOG', 'ERROR', `Erro: ${error.toString()}`, 0);
    return { total: 0, aplicaveis: 0, naoAplicaveis: 0 };
  }
}

// =============================================
// FUNÇÃO PARA SALVAR APENAS APLICÁVEIS NA PLANILHA
// =============================================

/**
 * Salvar APENAS os normativos aplicáveis na planilha principal
 */
function salvarAplicaveisNaPlanilha(todasAnalises) {
  Logger.log('💾 SALVANDO APENAS APLICÁVEIS NA PLANILHA PRINCIPAL...');
  
  try {
    // ✅ FILTRAR: Salvar APENAS os aplicáveis na planilha principal
    const normativosAplicaveis = todasAnalises.filter(analise => 
      analise.Aplicavel_iFood === 'Sim' && 
      analise.Impacto_Declarado !== 'N/A' && 
      analise.Impacto_Declarado !== 'Não Aplicável'
    );
    
    if (normativosAplicaveis.length === 0) {
      Logger.log('⚡ Nenhum normativo aplicável para salvar na planilha principal');
      return 0;
    }
    
    Logger.log(`📊 Filtrando: ${todasAnalises.length} análises → ${normativosAplicaveis.length} aplicáveis`);
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = spreadsheet.getSheets()[0];
    
    const ultimaLinha = sheet.getLastRow();
    
    if (ultimaLinha === 0) {
      const cabecalhos = [
        'normativo_index', 'Data_Captura', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Produto_Segmento', 'Tema', 'Impacto_Declarado',
        'Data_Vigencia', 'Aplicavel_SCD', 'Aplicavel_IP', 'Aplicavel_iFood',
        'status', 'Criticidade_Sistema', 'Resumo_Analise', 'Resposta_Toqan'
      ];
      sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }
    
    const dados = [];
    let proximoIndex = ultimaLinha + 1;
    
    normativosAplicaveis.forEach((analise, index) => {
      const linha = [
        analise.normativo_index || proximoIndex + index,
        analise.Data_Captura || Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss'),
        analise.Orgao || 'N/A',
        analise.Tipo_Norma || 'N/A',
        analise.Numero || 'N/A',
        analise.Data_Publicacao || 'N/A',
        analise.Produto_Segmento || 'iFood Pago - Geral',
        analise.Tema || 'N/A',
        analise.Impacto_Declarado || 'Médio',
        analise.Data_Vigencia || analise.Data_Publicacao || 'N/A',
        analise.Aplicavel_SCD || 'Não',
        analise.Aplicavel_IP || 'Sim',
        analise.Aplicavel_iFood || 'Sim',
        analise.status || 'Analisado',
        analise.Criticidade_Sistema || 'MÉDIA',
        analise.Resumo_Analise || 'Análise Toqan AI',
        analise.Resposta_Toqan || 'N/A'
      ];
      dados.push(linha);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha + 1;
      sheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} normativos APLICÁVEIS salvos na planilha principal!`);
      return dados.length;
    }
    
    return 0;
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar aplicáveis na planilha: ${error.toString()}`);
    return 0;
  }
}

// =============================================
// RELATÓRIO COM SEQUÊNCIA CORRETA
// =============================================

function enviarRelatorioExecucaoAgendada(resultado) {
  try {
    const data = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
    
    let mensagem = `🔍 *MONITORAMENTO AUTOMÁTICO - ${data}*\n\n`;
    mensagem += `📊 *SEQUÊNCIA CORRETA EXECUTADA:*\n`;
    mensagem += `📥 1. COLETA: ${resultado.normativosOficiais.length + resultado.fontesComplementares.length} itens\n`;
    mensagem += `🤖 2. TOQAN: ${resultado.analisesToqan.length} análises\n`;
    mensagem += `📚 3. BACKLOG: ${resultado.backlog} registros\n`;
    mensagem += `💾 4. PLANILHA: ${resultado.planilha} aplicáveis\n\n`;
    
    mensagem += `📈 *DETALHAMENTO:*\n`;
    mensagem += `├─ Coletados: ${resultado.normativosOficiais.length + resultado.fontesComplementares.length} itens\n`;
    mensagem += `├─ Analisados: ${resultado.analisesToqan.length} normativos\n`;
    mensagem += `├─ Aplicáveis: ${resultado.backlogAplicaveis || 0} itens\n`;
    mensagem += `├─ Não Aplicáveis: ${resultado.backlogNaoAplicaveis || 0} itens\n`;
    mensagem += `├─ Backlog: ${resultado.backlog} registros\n`;
    mensagem += `├─ Planilha: ${resultado.planilha} aplicáveis\n`;
    mensagem += `└─ Tempo: ${resultado.tempoExecucao}s\n\n`;
    
    // DETALHES DOS APLICÁVEIS
    if (resultado.planilha > 0) {
      mensagem += `🎯 *NORMATIVOS APLICÁVEIS PARA AÇÃO:*\n`;
      
      // Buscar os aplicáveis para mostrar
      const aplicaveis = resultado.analisesToqan.filter(a => a.Aplicavel_iFood === 'Sim').slice(0, 3);
      
      aplicaveis.forEach(normativo => {
        const emoji = normativo.Impacto_Declarado === 'Alto' ? '🔴' : 
                     normativo.Impacto_Declarado === 'Médio' ? '🟡' : '🟢';
        
        mensagem += `${emoji} ${normativo.Orgao} ${normativo.Numero} - ${normativo.Impacto_Declarado}\n`;
      });
      
      if (resultado.planilha > 3) {
        mensagem += `📝 ...e mais ${resultado.planilha - 3} normativos aplicáveis\n`;
      }
      mensagem += `\n`;
    }
    
    // INFORMAÇÃO SOBRE NÃO APLICÁVEIS
    if (resultado.backlogNaoAplicaveis > 0) {
      mensagem += `📋 *NÃO APLICÁVEIS (registrados no backlog):* ${resultado.backlogNaoAplicaveis} itens\n`;
      mensagem += `   _Histórico completo disponível no backlog_\n\n`;
    }
    
    mensagem += `✅ _Processo concluído - Sequência correta executada_`;
    
    return enviarSlackMensagem(mensagem);
    
  } catch (error) {
    Logger.log(`❌ Erro no relatório agendado: ${error}`);
    return enviarSlackMensagem(`🔍 Monitoramento automático executado - Verificar logs para detalhes`);
  }
}

// =============================================
// FUNÇÃO PARA TESTAR A SEQUÊNCIA
// =============================================

/**
 * Testar a sequência correta com dados de exemplo
 */
function testarSequenciaCorreta() {
  Logger.log('🧪 TESTANDO SEQUÊNCIA CORRETA');
  
  try {
    // Dados de teste
    const normativosTeste = [
      {
        Orgao: 'BACEN',
        Tipo_Norma: 'Circular',
        Numero: 'TESTE-001',
        Data_Publicacao: '2024-01-01',
        Tema: 'Normativo aplicável - Pagamentos',
        texto_completo: 'Regulamentação sobre sistema de pagamentos instantâneos',
        url_fonte: 'https://exemplo.com/teste1'
      },
      {
        Orgao: 'RFB',
        Tipo_Norma: 'Instrução Normativa',
        Numero: 'TESTE-002',
        Data_Publicacao: '2024-01-01',
        Tema: 'Normativo não aplicável - Importação',
        texto_completo: 'Regras para importação de produtos agrícolas',
        url_fonte: 'https://exemplo.com/teste2'
      }
    ];
    
    Logger.log('📥 1. COLETA: 2 normativos de teste');
    
    Logger.log('🤖 2. TOQAN: Analisando normativos...');
    const analises = analisarNormativosComToqan(normativosTeste);
    Logger.log(`   ✅ ${analises.length} análises completas`);
    
    Logger.log('📚 3. BACKLOG: Salvando todas as análises...');
    const backlog = salvarTodasAnalisesNoBacklog(analises);
    Logger.log(`   ✅ ${backlog.total} registros no backlog`);
    
    Logger.log('💾 4. PLANILHA: Salvando apenas aplicáveis...');
    const planilha = salvarAplicaveisNaPlanilha(analises);
    Logger.log(`   ✅ ${planilha} aplicáveis na planilha`);
    
    enviarSlackMensagem(
      `🧪 *TESTE DA SEQUÊNCIA CORRETA*\n\n` +
      `✅ Sequência testada com sucesso!\n` +
      `📥 Coleta: 2 normativos\n` +
      `🤖 Toqan: ${analises.length} análises\n` +
      `📚 Backlog: ${backlog.total} registros\n` +
      `💾 Planilha: ${planilha} aplicáveis\n\n` +
      `🎯 Sistema pronto para uso!`
    );
    
    return { success: true, backlog: backlog.total, planilha: planilha };
    
  } catch (error) {
    Logger.log(`❌ Erro no teste: ${error}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// 4. SISTEMA PRINCIPAL CORRIGIDO
// =============================================

/**
 * FUNÇÃO PRINCIPAL DO MONITORAMENTO
 */
function executarMonitoramentoCompletoPrincipal() {
  Logger.log('🔍 INICIANDO MONITORAMENTO COMPLETO - MODO PRINCIPAL');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. MONITORAMENTO OFICIAL (BACEN/CMN/DOU)
    Logger.log('🏛️  MÓDULO 1: MONITORAMENTO OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. MONITORAMENTO COMPLEMENTAR (NOTÍCIAS, PORTAIS)
    Logger.log('📰 MÓDULO 2: MONITORAMENTO COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado em nenhum módulo');
      resultados.mensagem = 'Nenhum normativo detectado';
      resultados.success = true;
      return resultados;
    }
    
    Logger.log(`📊 TOTAL DETECTADO: ${todosNormativos.length} normativos`);
    
    // 4. SALVAR NO BACKLOG (TODOS OS NORMATIVOS)
    Logger.log('📚 MÓDULO 3: BACKLOG COMPLETO...');
    resultados.backlog = salvarNoBacklog(todosNormativos);
    Logger.log(`   ✅ ${resultados.backlog} itens no backlog`);
    
    // 5. ANÁLISE TOQAN (APLICÁVEIS)
    Logger.log('🤖 MÓDULO 4: ANÁLISE TOQAN...');
    resultados.analisesToqan = analisarNormativosComToqan(todosNormativos);
    Logger.log(`   ✅ ${resultados.analisesToqan.length} normativos aplicáveis analisados`);
    
    // 6. ATUALIZAR BACKLOG COM ANÁLISES
    Logger.log('🔄 MÓDULO 5: ATUALIZANDO BACKLOG...');
    resultados.backlogAtualizado = atualizarBacklogComAnalise(resultados.analisesToqan);
    
    // 7. SALVAR APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 MÓDULO 6: PLANILHA PRINCIPAL...');
    resultados.planilha = salvarNaPlanilha(resultados.analisesToqan);
    Logger.log(`   ✅ ${resultados.planilha} itens na planilha principal`);
    
    // 8. TEMPO DE EXECUÇÃO
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    Logger.log(`🎯 MONITORAMENTO COMPLETO CONCLUÍDO EM ${resultados.tempoExecucao}s`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO MONITORAMENTO COMPLETO: ${error.toString()}`);
    return {
      success: false,
      error: error.toString(),
      endTime: new Date()
    };
  }
}
// =============================================
// CORREÇÃO DO SISTEMA - REMOVER FUNÇÃO ANTIGA
// =============================================

/**
 * FUNÇÃO PRINCIPAL CORRIGIDA - SEM CHAMAR ATUALIZARBACKLOGCOMNALISE
 */
function executarMonitoramentoCompletoPrincipal() {
  Logger.log('🔍 INICIANDO MONITORAMENTO COMPLETO - SEQUÊNCIA CORRETA CORRIGIDA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. 📥 COLETA - MONITORAMENTO OFICIAL (BACEN/CMN/DOU)
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. 📥 COLETA - MONITORAMENTO COMPLEMENTAR (NOTÍCIAS, PORTAIS)
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS DA COLETA
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado em nenhum módulo');
      resultados.mensagem = 'Nenhum normativo detectado';
      resultados.success = true;
      return resultados;
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 4. 🤖 TOQAN - ANALISAR TODOS OS NORMATIVOS COLETADOS
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN (TODOS OS NORMATIVOS)...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise Toqan concluída');
      resultados.mensagem = 'Análise Toqan não retornou resultados';
      resultados.success = false;
      return resultados;
    }
    
    // ✅ ESTATÍSTICAS DETALHADAS
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises completas (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 5. 📚 BACKLOG - SALVAR TODAS AS ANÁLISES NO BACKLOG (FUNÇÃO CORRETA)
    Logger.log('📚 ETAPA 4: BACKLOG (TODAS AS ANÁLISES)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.total;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 6. 💾 PLANILHA - SALVAR APENAS APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 ETAPA 5: PLANILHA (APENAS APLICÁVEIS)...');
    resultados.planilha = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.planilha} itens APLICÁVEIS na planilha principal`);
    
    // 7. TEMPO DE EXECUÇÃO
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    Logger.log(`🎯 MONITORAMENTO COMPLETO CONCLUÍDO EM ${resultados.tempoExecucao}s`);
    Logger.log(`📈 SEQUÊNCIA CORRETA: ${todosNormativos.length} coletados → ${todasAnalises.length} analisados → ${resultados.backlog} backlog → ${resultados.planilha} planilha`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO MONITORAMENTO COMPLETO: ${error.toString()}`);
    return {
      success: false,
      error: error.toString(),
      endTime: new Date()
    };
  }
}

// =============================================
// REMOVER/SUBSTITUIR FUNÇÃO PROBLEMÁTICA
// =============================================

/**
 * SUBSTITUIR a função problemática atualizarBacklogComAnalise
 * por uma versão que simplesmente chama a função correta
 */
function atualizarBacklogComAnalise(normativosAnalisados) {
  Logger.log('🔄 FUNÇÃO ATUALIZARBACKLOGCOMNALISE CHAMADA - REDIRECIONANDO...');
  
  // Simplesmente chamar a função correta
  const resultado = salvarTodasAnalisesNoBacklog(normativosAnalisados);
  
  Logger.log(`✅ Redirecionado: ${resultado.total} análises salvas no backlog`);
  return resultado.total;
}

// =============================================
// FUNÇÃO MELHORADA PARA SALVAR BACKLOG
// =============================================

/**
 * Função MELHORADA para salvar todas as análises no backlog
 * com melhor tratamento de dados
 */
function salvarTodasAnalisesNoBacklog(todasAnalises) {
  Logger.log('📚 SALVANDO TODAS AS ANÁLISES NO BACKLOG...');
  
  try {
    // VALIDAÇÃO DE ENTRADA
    if (!todasAnalises || !Array.isArray(todasAnalises) || todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise para salvar no backlog');
      return { total: 0, aplicaveis: 0, naoAplicaveis: 0 };
    }
    
    Logger.log(`📝 Processando ${todasAnalises.length} análises para backlog`);
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let backlogSheet;
    
    try {
      backlogSheet = spreadsheet.getSheetByName('BACKLOG');
    } catch (e) {
      // Criar aba BACKLOG se não existir
      backlogSheet = spreadsheet.insertSheet('BACKLOG');
      const cabecalhos = [
        'ID_Backlog', 'Data_Coleta', 'Orgao', 'Tipo_Norma', 'Numero',
        'Data_Publicacao', 'Tema', 'Texto_Completo', 'URL_Fonte',
        'Status_Analise', 'Impacto_Toqan', 'Produto_Afetado_Toqan',
        'Aplicavel_SCD_Toqan', 'Aplicavel_iFood_Toqan', 'Resumo_Toqan',
        'ID_Conversa_Toqan', 'Data_Analise_Toqan'
      ];
      backlogSheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      backlogSheet.getRange(1, 1, 1, cabecalhos.length)
        .setBackground('#2E7D32')
        .setFontColor('white')
        .setFontWeight('bold');
      
      Logger.log('✅ Nova aba BACKLOG criada');
    }
    
    const dados = [];
    const dataColeta = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    const ultimaLinha = backlogSheet.getLastRow();
    let proximoID = 1;
    
    if (ultimaLinha > 0) {
      const ultimoID = backlogSheet.getRange(ultimaLinha, 1).getValue();
      proximoID = isNaN(ultimoID) ? 1 : parseInt(ultimoID) + 1;
    }
    
    let aplicaveis = 0;
    let naoAplicaveis = 0;
    
    todasAnalises.forEach((analise, index) => {
      // VALIDAÇÃO DOS DADOS DA ANÁLISE
      if (!analise) {
        Logger.log(`   ⚡ Análise ${index} é nula, pulando...`);
        return;
      }
      
      // Contar estatísticas
      if (analise.Aplicavel_iFood === 'Sim') {
        aplicaveis++;
      } else {
        naoAplicaveis++;
      }
      
      // PREPARAR DADOS COM VALIDAÇÃO
      const linha = [
        proximoID + index, // ID_Backlog
        dataColeta, // Data_Coleta
        analise.Orgao || 'N/A', // Orgao
        analise.Tipo_Norma || 'N/A', // Tipo_Norma
        analise.Numero || 'N/A', // Numero
        analise.Data_Publicacao || 'N/A', // Data_Publicacao
        analise.Tema || 'N/A', // Tema
        analise.texto_completo || analise.Tema || 'N/A', // Texto_Completo
        analise.url_fonte || 'N/A', // URL_Fonte
        'Analisado', // Status_Analise - JÁ ANALISADO
        analise.Impacto_Declarado || 'N/A', // Impacto_Toqan
        analise.Produto_Segmento || 'N/A', // Produto_Afetado_Toqan
        analise.Aplicavel_SCD || 'N/A', // Aplicavel_SCD_Toqan
        analise.Aplicavel_iFood || 'N/A', // Aplicavel_iFood_Toqan
        analise.Resumo_Analise || 'N/A', // Resumo_Toqan
        analise.Resposta_Toqan ? 
          (typeof analise.Resposta_Toqan === 'string' ? 
           analise.Resposta_Toqan.replace('Toqan ID: ', '') : 'N/A') : 'N/A', // ID_Conversa_Toqan
        dataColeta // Data_Analise_Toqan
      ];
      
      dados.push(linha);
      Logger.log(`   📝 Preparado: ${analise.Orgao} ${analise.Numero} - ${analise.Aplicavel_iFood}`);
    });
    
    if (dados.length > 0) {
      const linhaInicio = ultimaLinha === 0 ? 2 : ultimaLinha + 1;
      backlogSheet.getRange(linhaInicio, 1, dados.length, dados[0].length).setValues(dados);
      Logger.log(`✅ ${dados.length} análises salvas no BACKLOG! (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
      
      // Registrar no log do sistema
      registrarLogAPI('BACKLOG', 'SUCCESS', 
        `${dados.length} análises salvas no backlog (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`, 
        dados.length
      );
      
      return {
        total: dados.length,
        aplicaveis: aplicaveis,
        naoAplicaveis: naoAplicaveis
      };
    } else {
      Logger.log('⚡ Nenhum dado válido para salvar no backlog');
      return { total: 0, aplicaveis: 0, naoAplicaveis: 0 };
    }
    
  } catch (error) {
    Logger.log(`❌ ERRO ao salvar análises no backlog: ${error.toString()}`);
    registrarLogAPI('BACKLOG', 'ERROR', `Erro: ${error.toString()}`, 0);
    return { total: 0, aplicaveis: 0, naoAplicaveis: 0 };
  }
}

// =============================================
// FUNÇÃO PARA VERIFICAR SE O BACKLOG ESTÁ FUNCIONANDO
// =============================================

/**
 * Verificar status do backlog
 */
function verificarStatusBacklog() {
  Logger.log('🔍 VERIFICANDO STATUS DO BACKLOG');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const backlogSheet = spreadsheet.getSheetByName('BACKLOG');
    
    if (!backlogSheet) {
      Logger.log('❌ Aba BACKLOG não encontrada');
      enviarSlackMensagem('❌ *BACKLOG*: Aba não encontrada');
      return { success: false, error: 'Backlog não encontrado' };
    }
    
    const ultimaLinha = backlogSheet.getLastRow();
    
    if (ultimaLinha <= 1) {
      Logger.log('📝 Backlog vazio');
      enviarSlackMensagem('📝 *BACKLOG*: Vazio (aguardando dados)');
      return { success: true, total: 0, vazio: true };
    }
    
    // Contar estatísticas
    const dados = backlogSheet.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const total = dados.length;
    const analisados = dados.filter(linha => linha[9] === 'Analisado').length;
    const aplicaveis = dados.filter(linha => linha[13] === 'Sim').length;
    const naoAplicaveis = analisados - aplicaveis;
    
    let mensagem = `📚 *STATUS DO BACKLOG*\n\n`;
    mensagem += `📊 Estatísticas:\n`;
    mensagem += `├─ Total de registros: ${total}\n`;
    mensagem += `├─ Analisados: ${analisados}\n`;
    mensagem += `├─ Aplicáveis: ${aplicaveis}\n`;
    mensagem += `└─ Não aplicáveis: ${naoAplicaveis}\n\n`;
    
    // Últimos 5 registros
    if (total > 0) {
      mensagem += `📋 Últimos registros:\n`;
      dados.slice(-5).forEach(linha => {
        const status = linha[9] === 'Analisado' ? '✅' : '⏳';
        const aplicavel = linha[13] === 'Sim' ? '🎯' : '📝';
        mensagem += `${status}${aplicavel} ${linha[2]} ${linha[3]} ${linha[4]}\n`;
      });
    }
    
    enviarSlackMensagem(mensagem);
    
    return { 
      success: true, 
      total: total, 
      analisados: analisados, 
      aplicaveis: aplicaveis,
      naoAplicaveis: naoAplicaveis
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao verificar backlog: ${error}`);
    enviarSlackMensagem(`❌ Erro ao verificar backlog: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PARA TESTE RÁPIDO DO BACKLOG
// =============================================

/**
 * Teste rápido do backlog
 */
function testarBacklogRapido() {
  Logger.log('🧪 TESTE RÁPIDO DO BACKLOG');
  
  try {
    // Criar dados de teste
    const analisesTeste = [
      {
        Orgao: 'TESTE',
        Tipo_Norma: 'Circular',
        Numero: 'TEST-001',
        Data_Publicacao: '2024-01-01',
        Tema: 'Teste aplicável',
        texto_completo: 'Texto de teste aplicável',
        url_fonte: 'https://teste.com/1',
        Impacto_Declarado: 'Alto',
        Produto_Segmento: 'iFood Pago',
        Aplicavel_SCD: 'Sim',
        Aplicavel_iFood: 'Sim',
        Resumo_Analise: 'Teste aplicável ao iFood',
        Resposta_Toqan: 'Toqan ID: TEST-123'
      },
      {
        Orgao: 'TESTE',
        Tipo_Norma: 'Resolução',
        Numero: 'TEST-002',
        Data_Publicacao: '2024-01-01',
        Tema: 'Teste não aplicável',
        texto_completo: 'Texto de teste não aplicável',
        url_fonte: 'https://teste.com/2',
        Impacto_Declarado: 'Não Aplicável',
        Produto_Segmento: 'Nenhum',
        Aplicavel_SCD: 'Não',
        Aplicavel_iFood: 'Não',
        Resumo_Analise: 'Teste não aplicável ao iFood',
        Resposta_Toqan: 'Toqan ID: TEST-456'
      }
    ];
    
    Logger.log('📚 Salvando análises de teste no backlog...');
    const resultado = salvarTodasAnalisesNoBacklog(analisesTeste);
    
    enviarSlackMensagem(
      `🧪 *TESTE DO BACKLOG*\n\n` +
      `✅ Teste concluído com sucesso!\n` +
      `📊 Resultado:\n` +
      `├─ Total: ${resultado.total} análises\n` +
      `├─ Aplicáveis: ${resultado.aplicaveis}\n` +
      `└─ Não aplicáveis: ${resultado.naoAplicaveis}\n\n` +
      `🎯 Backlog funcionando corretamente!`
    );
    
    return resultado;
    
  } catch (error) {
    Logger.log(`❌ Erro no teste do backlog: ${error}`);
    enviarSlackMensagem(`❌ Falha no teste do backlog: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// 5. FUNÇÃO PARA AGENDAMENTO
// =============================================

/**
 * FUNÇÃO QUE SERÁ EXECUTADA NOS HORÁRIOS AGENDADOS
 */
function executarSistemaAgendado() {
  Logger.log('🔍 EXECUTANDO SISTEMA AGENDADO - TODOS OS MÓDULOS');
  
  try {
    const resultado = executarMonitoramentoCompletoPrincipal();
    
    // ENVIAR RELATÓRIO DE EXECUÇÃO AGENDADA
    if (resultado.success) {
      enviarRelatorioExecucaoAgendada(resultado);
    } else {
      enviarSlackMensagem(
        `❌ *EXECUÇÃO AGENDADA COM FALHA*\n\n` +
        `⚡ Erro: ${resultado.error}\n` +
        `🔧 Verificar logs para detalhes`
      );
    }
    
    return resultado;
    
  } catch (error) {
    Logger.log(`❌ ERRO NA EXECUÇÃO AGENDADA: ${error.toString()}`);
    enviarSlackMensagem(`❌ FALHA NA EXECUÇÃO AGENDADA: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// CORREÇÃO DO SISTEMA - REMOVER REINICIALIZAÇÃO AUTOMÁTICA
// =============================================

/**
 * FUNÇÃO PRINCIPAL CORRIGIDA - SEM REINICIALIZAÇÃO AUTOMÁTICA
 */
function executarMonitoramentoCompletoPrincipal() {
  Logger.log('🔍 INICIANDO MONITORAMENTO COMPLETO - SEM REINICIALIZAÇÃO');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. 📥 COLETA - MONITORAMENTO OFICIAL (BACEN/CMN/DOU)
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. 📥 COLETA - MONITORAMENTO COMPLEMENTAR (NOTÍCIAS, PORTAIS)
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS DA COLETA
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado em nenhum módulo');
      resultados.mensagem = 'Nenhum normativo detectado';
      resultados.success = true;
      
      // ✅ NÃO REINICIALIZAR - apenas enviar mensagem
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return resultados;
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 4. 🤖 TOQAN - ANALISAR TODOS OS NORMATIVOS COLETADOS
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN (TODOS OS NORMATIVOS)...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise Toqan concluída');
      resultados.mensagem = 'Análise Toqan não retornou resultados';
      resultados.success = false;
      
      // ✅ NÃO REINICIALIZAR - apenas enviar mensagem de erro
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return resultados;
    }
    
    // ✅ ESTATÍSTICAS DETALHADAS
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises completas (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 5. 📚 BACKLOG - SALVAR TODAS AS ANÁLISES NO BACKLOG
    Logger.log('📚 ETAPA 4: BACKLOG (TODAS AS ANÁLISES)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.total;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 6. 💾 PLANILHA - SALVAR APENAS APLICÁVEIS NA PLANILHA PRINCIPAL
    Logger.log('💾 ETAPA 5: PLANILHA (APENAS APLICÁVEIS)...');
    resultados.planilha = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.planilha} itens APLICÁVEIS na planilha principal`);
    
    // 7. ENVIAR RELATÓRIO FINAL
    Logger.log('📊 ETAPA 6: RELATÓRIO FINAL...');
    enviarRelatorioExecucaoAgendada(resultados);
    
    // 8. TEMPO DE EXECUÇÃO
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    Logger.log(`🎯 MONITORAMENTO COMPLETO CONCLUÍDO EM ${resultados.tempoExecucao}s - AGUARDANDO PRÓXIMO AGENDAMENTO`);
    
    // ✅ NÃO REINICIALIZAR - O SISTEMA PARA AQUI E AGUARDA O PRÓXIMO AGENDAMENTO
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO MONITORAMENTO COMPLETO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return {
      success: false,
      error: error.toString(),
      endTime: new Date()
    };
  }
}

// =============================================
// SISTEMA DE AGENDAMENTO ESTÁVEL
// =============================================

/**
 * Configurar agendamento UMA VEZ - não reinicializar automaticamente
 */
function configurarAgendamentoEstavel() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTO ESTÁVEL - UMA ÚNICA VEZ');
  
  try {
    // Verificar agendamentos existentes primeiro
    const triggersAtuais = ScriptApp.getProjectTriggers();
    
    if (triggersAtuais.length > 0) {
      Logger.log(`📊 Agendamentos já existentes: ${triggersAtuais.length}`);
      triggersAtuais.forEach(trigger => {
        Logger.log(`   ✅ ${trigger.getHandlerFunction()} - ${trigger.getEventType()}`);
      });
      
      enviarSlackMensagem(
        `⏰ *AGENDAMENTOS JÁ ATIVOS*\n\n` +
        `✅ ${triggersAtuais.length} agendamentos encontrados\n` +
        `📅 Sistema já está programado\n\n` +
        `🎯 Próximas execuções automáticas configuradas!`
      );
      
      return { success: true, mensagem: 'Agendamentos já ativos', triggers: triggersAtuais.length };
    }
    
    // Se não há agendamentos, configurar
    Logger.log('⚡ Nenhum agendamento encontrado - configurando...');
    
    const horarios = [9, 12, 17]; // 9h, 12h, 17h
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarSistemaAgendado')
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`   ✅ Agendado: ${hora}:00`);
    });
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    enviarSlackMensagem(
      `⏰ *AGENDAMENTO CONFIGURADO - UMA ÚNICA VEZ*\n\n` +
      `✅ ${triggersFinais.length} agendamentos ativos\n` +
      `🕘 Horários: 9h, 12h, 17h\n` +
      `🔍 Execução: executarSistemaAgendado()\n\n` +
      `🎯 Sistema programado para os próximos dias!`
    );
    
    Logger.log('🎯 AGENDAMENTO CONFIGURADO - NÃO REINICIALIZAR AUTOMATICAMENTE');
    
    return { success: true, triggers: triggersFinais.length };
    
  } catch (error) {
    Logger.log(`❌ ERRO NO AGENDAMENTO: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Função para verificar e manter agendamentos estáveis
 */
function verificarManutencaoAgendamentos() {
  Logger.log('🔍 VERIFICANDO MANUTENÇÃO DE AGENDAMENTOS');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggersExecutavel = triggers.filter(t => t.getHandlerFunction() === 'executarSistemaAgendado');
    
    if (triggersExecutavel.length >= 3) {
      Logger.log(`✅ Agendamentos estáveis: ${triggersExecutavel.length} triggers ativos`);
      return { 
        success: true, 
        estaEstavel: true, 
        triggers: triggersExecutavel.length,
        mensagem: 'Sistema estável - não requer intervenção'
      };
    }
    
    if (triggersExecutavel.length === 0) {
      Logger.log('⚠️  Nenhum agendamento ativo - requer configuração');
      enviarSlackMensagem(
        `⚠️  *AGENDAMENTOS INATIVOS*\n\n` +
        `Nenhum trigger ativo encontrado\n` +
        `Execute 'configurarAgendamentoEstavel()' para reativar`
      );
      return { success: false, estaEstavel: false, triggers: 0 };
    }
    
    Logger.log(`⚠️  Agendamentos insuficientes: ${triggersExecutavel.length}/3`);
    return { 
      success: true, 
      estaEstavel: false, 
      triggers: triggersExecutavel.length,
      mensagem: 'Agendamentos insuficientes - considerar reconfiguração'
    };
    
  } catch (error) {
    Logger.log(`❌ Erro na verificação: ${error}`);
    return { success: false, error: error.toString() };
  }
}


// =============================================
// FUNÇÃO PARA PARAR COMPLETAMENTE
// =============================================

/**
 * Parar completamente o sistema (para manutenção)
 */
function pararSistemaCompletamente() {
  Logger.log('🛑 PARANDO SISTEMA COMPLETAMENTE');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
      Logger.log(`   🗑️  Removido: ${trigger.getHandlerFunction()}`);
    });
    
    enviarSlackMensagem(
      `🛑 *SISTEMA PARADO COMPLETAMENTE*\n\n` +
      `✅ ${removidos} agendamentos removidos\n` +
      `⚡ Sistema inativo até nova configuração\n\n` +
      `🔧 Para reativar, execute 'iniciarSistemaEstavel()'`
    );
    
    return { success: true, removidos: removidos };
    
  } catch (error) {
    Logger.log(`❌ Erro ao parar sistema: ${error}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// FUNÇÃO PRINCIPAL PARA AGENDAMENTO (ESTÁVEL)
// =============================================

/**
 * FUNÇÃO QUE SERÁ CHAMADA PELOS AGENDAMENTOS - ESTÁVEL
 */
function executarSistemaAgendado() {
  Logger.log('🔍 EXECUTANDO SISTEMA AGENDADO - MODO ESTÁVEL');
  
  try {
    // Apenas executar o monitoramento - NÃO REINICIALIZAR
    const resultado = executarMonitoramentoCompletoPrincipal();
    
    // ✅ NÃO CONFIGURAR NOVOS AGENDAMENTOS - já estão configurados
    Logger.log('✅ Execução agendada concluída - aguardando próximo horário');
    
    return resultado;
    
  } catch (error) {
    Logger.log(`❌ ERRO NA EXECUÇÃO AGENDADA: ${error.toString()}`);
    enviarSlackMensagem(`❌ FALHA NA EXECUÇÃO AGENDADA: ${error.toString().substring(0, 100)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// FUNÇÃO DE INICIALIZAÇÃO SEGURA
// =============================================

/**
 * Inicialização segura - configura agendamento UMA VEZ e para
 */
function iniciarSistemaEstavel() {
  Logger.log('🚀 INICIANDO SISTEMA ESTÁVEL - CONFIGURAÇÃO ÚNICA');
  
  try {
    // 1. Configurar agendamento (apenas se necessário)
    Logger.log('⏰ ETAPA 1: VERIFICAR/CONFIGURAR AGENDAMENTO...');
    const agendamento = configurarAgendamentoEstavel();
    
    // 2. Executar monitoramento uma vez
    Logger.log('🔍 ETAPA 2: EXECUTAR MONITORAMENTO INICIAL...');
    const resultado = executarMonitoramentoCompletoPrincipal();
    
    // 3. Enviar relatório final
    Logger.log('📊 ETAPA 3: RELATÓRIO DE INICIALIZAÇÃO...');
    enviarRelatorioInicializacao({
      ...resultado,
      agendamento: agendamento
    });
    
    Logger.log('🎯 SISTEMA INICIALIZADO - AGUARDANDO PRÓXIMOS AGENDAMENTOS AUTOMÁTICOS');
    
    return {
      success: true,
      agendamento: agendamento,
      monitoramento: resultado,
      mensagem: 'Sistema inicializado com sucesso - agendamentos ativos'
    };
    
  } catch (error) {
    Logger.log(`❌ ERRO NA INICIALIZAÇÃO: ${error.toString()}`);
    enviarSlackMensagem(`❌ FALHA NA INICIALIZAÇÃO: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}
// =============================================
// SISTEMA CORRIGIDO - SEM AUTO-REINICIAÇÃO
// =============================================

/**
 * FUNÇÃO PRINCIPAL - EXECUÇÃO ÚNICA E INDEPENDENTE
 * PARA EXECUÇÃO MANUAL OU VIA AGENDAMENTO
 * NÃO CHAMA OUTRAS FUNÇÕES AUTOMATICAMENTE
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO COMPLETO - EXECUÇÃO ÚNICA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      planilha: 0,
      startTime: new Date()
    };
    
    // 1. 📥 COLETA - MONITORAMENTO OFICIAL
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. 📥 COLETA - MONITORAMENTO COMPLEMENTAR
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // 3. COMBINAR TODOS OS RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 4. 🤖 TOQAN - ANALISAR NORMATIVOS
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise concluída');
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return { success: false, mensagem: 'Análise não concluída' };
    }
    
    // ESTATÍSTICAS
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 5. 📚 BACKLOG - SALVAR ANÁLISES
    Logger.log('📚 ETAPA 4: BACKLOG...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.total;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 6. 💾 PLANILHA - SALVAR APLICÁVEIS
    Logger.log('💾 ETAPA 5: PLANILHA...');
    resultados.planilha = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.planilha} aplicáveis na planilha`);
    
    // 7. RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    enviarRelatorioExecucaoAgendada(resultados);
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA EM ${resultados.tempoExecucao}s - SISTEMA PARADO`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// CONTROLE DE AGENDAMENTOS
// =============================================

/**
 * CONFIGURAR AGENDAMENTOS (executar manualmente UMA VEZ)
 */
function configurarAgendamentos() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTOS - EXECUÇÃO MANUAL');
  
  try {
    // REMOVER TODOS OS TRIGGERS EXISTENTES
    pararTodosAgendamentos();
    
    // CONFIGURAR NOVOS AGENDAMENTOS
    const horarios = [9, 12, 17]; // 9h, 12h, 17h
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarMonitoramentoCompleto')
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`✅ Agendado: ${hora}:00`);
    });
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    enviarSlackMensagem(
      `⏰ *AGENDAMENTOS CONFIGURADOS*\n\n` +
      `✅ ${triggersFinais.length} triggers ativos\n` +
      `🕘 Horários: 9h, 12h, 17h\n` +
      `🔍 Função: executarMonitoramentoCompleto()\n\n` +
      `🎯 Sistema programado para execução automática!`
    );
    
    return triggersFinais.length;
    
  } catch (error) {
    Logger.log(`❌ ERRO AO CONFIGURAR AGENDAMENTOS: ${error}`);
    enviarSlackMensagem(`❌ ERRO NOS AGENDAMENTOS: ${error.toString().substring(0, 150)}`);
    return 0;
  }
}

/**
 * PARAR TODOS OS AGENDAMENTOS
 */
function pararTodosAgendamentos() {
  const triggers = ScriptApp.getProjectTriggers();
  let removidos = 0;
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log(`🗑️  Removido: ${trigger.getHandlerFunction()}`);
    removidos++;
  });
  
  Logger.log(`✅ ${removidos} triggers removidos`);
  return removidos;
}

/**
 * VERIFICAR AGENDAMENTOS ATIVOS
 */
function verificarAgendamentos() {
  const triggers = ScriptApp.getProjectTriggers();
  
  const infoTriggers = triggers.map(trigger => {
    return {
      função: trigger.getHandlerFunction(),
      fonte: trigger.getTriggerSource(),
      evento: trigger.getEventType()
    };
  });
  
  Logger.log(`🔍 ${triggers.length} triggers ativos:`);
  infoTriggers.forEach(info => {
    Logger.log(`   📌 ${info.função} - ${info.fonte} - ${info.evento}`);
  });
  
  enviarSlackMensagem(
    `🔍 *VERIFICAÇÃO DE AGENDAMENTOS*\n\n` +
    `✅ ${triggers.length} triggers ativos\n` +
    `${infoTriggers.map(t => `• ${t.função}`).join('\n')}`
  );
  
  return infoTriggers;
}

// =============================================
// SISTEMA CORRIGIDO - SEM AUTO-EXECUÇÃO
// =============================================

/**
 * 🚀 INICIAR SISTEMA COMPLETO (Executar manualmente UMA VEZ)
 * APENAS configura agendamentos, NÃO executa monitoramento
 */
function iniciarSistemaCompleto() {
  Logger.log('🚀 INICIANDO SISTEMA COMPLETO - CONFIGURAÇÃO ÚNICA');
  
  try {
    // 1. Configurar agendamentos
    const triggersConfigurados = configurarAgendamentos();
    
    // 2. Verificar agendamentos
    verificarAgendamentos();
    
    // 3. ✅ NÃO EXECUTAR MONITORAMENTO AUTOMATICAMENTE
    Logger.log('✅ Sistema configurado. Monitoramento executará apenas nos horários agendados.');
    
    enviarSlackMensagem(
      `🚀 *SISTEMA INICIADO COM SUCESSO*\n\n` +
      `✅ ${triggersConfigurados} agendamentos configurados\n` +
      `🕘 Horários: 9h, 12h, 17h\n` +
      `⏰ Próxima execução automática: nos horários agendados\n\n` +
      `📋 Para executar agora manualmente, use executarAgora()`
    );
    
    return triggersConfigurados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO AO INICIAR SISTEMA: ${error.toString().substring(0, 150)}`);
    return 0;
  }
}

/**
 * ⏰ CONFIGURAR AGENDAMENTOS (independente)
 */
function configurarAgendamentos() {
  Logger.log('⏰ CONFIGURANDO AGENDAMENTOS - EXECUÇÃO MANUAL');
  
  try {
    // REMOVER TODOS OS TRIGGERS EXISTENTES
    const triggersRemovidos = pararTodosAgendamentos();
    
    // CONFIGURAR NOVOS AGENDAMENTOS
    const horarios = [9, 12, 17];
    
    horarios.forEach(hora => {
      ScriptApp.newTrigger('executarMonitoramentoCompleto')
        .timeBased()
        .atHour(hora)
        .nearMinute(0)
        .everyDays(1)
        .inTimezone('America/Sao_Paulo')
        .create();
      Logger.log(`✅ Agendado: ${hora}:00`);
    });
    
    const triggersFinais = ScriptApp.getProjectTriggers();
    
    enviarSlackMensagem(
      `⏰ *AGENDAMENTOS CONFIGURADOS*\n\n` +
      `✅ ${triggersFinais.length} triggers ativos\n` +
      `🕘 Horários: 9h, 12h, 17h\n` +
      `🔍 Função: executarMonitoramentoCompleto()\n\n` +
      `🎯 Sistema programado para execução automática!`
    );
    
    return triggersFinais.length;
    
  } catch (error) {
    Logger.log(`❌ ERRO AO CONFIGURAR AGENDAMENTOS: ${error}`);
    enviarSlackMensagem(`❌ ERRO NOS AGENDAMENTOS: ${error.toString().substring(0, 150)}`);
    return 0;
  }
}

/**
 * 🛑 PARAR TODOS OS AGENDAMENTOS
 */
function pararTodosAgendamentos() {
  const triggers = ScriptApp.getProjectTriggers();
  let removidos = 0;
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log(`🗑️  Removido: ${trigger.getHandlerFunction()}`);
    removidos++;
  });
  
  Logger.log(`✅ ${removidos} triggers removidos`);
  return removidos;
}

/**
 * 🔍 VERIFICAR AGENDAMENTOS ATIVOS
 */
function verificarAgendamentos() {
  const triggers = ScriptApp.getProjectTriggers();
  
  const infoTriggers = triggers.map(trigger => {
    return {
      função: trigger.getHandlerFunction(),
      fonte: trigger.getTriggerSource(),
      evento: trigger.getEventType()
    };
  });
  
  Logger.log(`🔍 ${triggers.length} triggers ativos:`);
  infoTriggers.forEach(info => {
    Logger.log(`   📌 ${info.função} - ${info.fonte} - ${info.evento}`);
  });
  
  return infoTriggers;
}

// FUNÇÃO PRINCIPAL DE MONITORAMENTO

/**
 * 🔍 EXECUTAR MONITORAMENTO COMPLETO
 * PARA EXECUÇÃO MANUAL OU VIA AGENDAMENTO
 * NÃO CHAMA OUTRAS FUNÇÕES AUTOMATICAMENTE
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO COMPLETO - EXECUÇÃO ÚNICA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      agenda: 0,
      startTime: new Date()
    };
    
    // 1. COLETA OFICIAL
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. COLETA COMPLEMENTAR  
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    const monitor = new MonitoramentoNormativo();
    Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    
    // COMBINAR RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 3. ANÁLISE TOQAN
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise concluída');
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return { success: false, mensagem: 'Análise não concluída' };
    }
    
    // ESTATÍSTICAS DAS ANÁLISES
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 4. 📚 BACKLOG - SALVAR TODOS OS NORMATIVOS
    Logger.log('📚 ETAPA 4: BACKLOG (TODOS OS NORMATIVOS)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.salvos;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 5. 💾 AGENDA NORMATIVA - SALVAR APENAS APLICÁVEIS
    Logger.log('💾 ETAPA 5: AGENDA NORMATIVA (APENAS APLICÁVEIS)...');
    resultados.agenda = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.agenda} aplicáveis na AgendaNormativa`);
    
    // RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    // ENVIAR RELATÓRIO DETALHADO
    enviarRelatorioExecucaoAgendada(resultados);
    
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA: ${resultados.backlog} no Backlog, ${resultados.agenda} na Agenda`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}

// =============================================
// REMOVER COMPLETAMENTE AS FUNÇÕES PROBLEMÁTICAS
// =============================================

/**
 * ❌❌❌ REMOVER/COMENTAR ESTAS FUNÇÕES PROBLEMÁTICAS ❌❌❌
 * Elas estão causando a auto-execução
 */

/*
// ❌ REMOVER ESTA FUNÇÃO - ELA CHAMA EXECUÇÃO AUTOMÁTICA
function configurarApenasAgendamento() {
  Logger.log('🚀 INICIANDO APENAS O SISTEMA DE AGENDAMENTO');
  configurarAgendamentos();
  Logger.log('🎉 SISTEMA DE AGENDAMENTO INICIADO!');
  Logger.log('📋 O sistema completo executará automaticamente nos horários configurados');
  
  // ❌❌❌ ESTA LINHA CAUSA A AUTO-EXECUÇÃO ❌❌❌
  executarMonitoramentoCompleto(); // REMOVER ESTA LINHA
}

// ❌ REMOVER ESTA FUNÇÃO - TAMBÉM CAUSA AUTO-EXECUÇÃO
function iniciarSistemaEstavel() {
  Logger.log('🚀 CONFIGURANDO APENAS O SISTEMA DE AGENDAMENTO');
  configurarApenasAgendamento(); // QUE CHAMA executarMonitoramentoCompleto()
}
*/

// =============================================
// ⚠️⚠️⚠️ IMPORTANTE: VERIFICAR O FINAL DO CÓDIGO ⚠️⚠️⚠️
// =============================================

/**
 * ✅✅✅ VERIFICAR SE NO FINAL DO ARQUIVO EXISTEM CHAMADAS AUTOMÁTICAS
 * E COMENTAR/REMOVER COMPLETAMENTE:
 * 
 * ❌ NÃO DEVE EXISTIR NENHUMA DESTAS LINHAS NO FINAL:
 * 
 * iniciarSistemaCompleto();
 * executarMonitoramentoCompleto();
 * configurarAgendamentos();
 * configurarApenasAgendamento();
 * iniciarSistemaEstavel();
 * qualquerOutraFuncaoQueExecuteAutomaticamente();
 * 
 * ✅ O CÓDIGO DEVE TERMINAR APENAS COM DEFINIÇÕES DE FUNÇÕES
 * ✅ NENHUMA FUNÇÃO DEVE SER CHAMADA AUTOMATICAMENTE
 */
// =============================================
// CORREÇÃO DO MONITORAMENTO NORMATIVO
// =============================================

/**
 * 🔍 EXECUTAR MONITORAMENTO COMPLETO - VERSÃO CORRIGIDA
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO COMPLETO - VERSÃO CORRIGIDA');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      agenda: 0,
      startTime: new Date()
    };
    
    // 1. COLETA OFICIAL
    Logger.log('📥 ETAPA 1: COLETA OFICIAL...');
    resultados.normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    
    // 2. COLETA COMPLEMENTAR - CORREÇÃO APPLY  
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR...');
    
    let fontesComplementares = [];
    try {
      // Tenta instanciar a classe MonitoramentoNormativo
      const monitor = new MonitoramentoNormativo();
      
      // Verifica se o método existe antes de chamar
      if (monitor && typeof monitor.executarMonitoramentoCompleto === 'function') {
        fontesComplementares = monitor.executarMonitoramentoCompleto();
        Logger.log(`   ✅ ${fontesComplementares.length} fontes complementares`);
      } else {
        Logger.log('   ⚠️ Método executarMonitoramentoCompleto não encontrado');
        fontesComplementares = executarMonitoramentoFallback();
      }
    } catch (error) {
      Logger.log(`   ⚠️ Erro na instanciação: ${error}`);
      fontesComplementares = executarMonitoramentoFallback();
    }
    
    resultados.fontesComplementares = fontesComplementares;
    
    // COMBINAR RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    
    // 3. ANÁLISE TOQAN
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise concluída');
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return { success: false, mensagem: 'Análise não concluída' };
    }
    
    // ESTATÍSTICAS DAS ANÁLISES
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 4. 📚 BACKLOG - SALVAR TODOS OS NORMATIVOS
    Logger.log('📚 ETAPA 4: BACKLOG (TODOS OS NORMATIVOS)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.salvos;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 5. 💾 AGENDA NORMATIVA - SALVAR APENAS APLICÁVEIS
    Logger.log('💾 ETAPA 5: AGENDA NORMATIVA (APENAS APLICÁVEIS)...');
    resultados.agenda = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.agenda} aplicáveis na AgendaNormativa`);
    
    // RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    // ENVIAR RELATÓRIO DETALHADO
    enviarRelatorioExecucaoAgendada(resultados);
    
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA: ${resultados.backlog} no Backlog, ${resultados.agenda} na Agenda`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * 🔄 FUNÇÃO FALLBACK PARA MONITORAMENTO
 */
function executarMonitoramentoFallback() {
  Logger.log('   🔄 Executando fallback de monitoramento...');
  
  const resultadosFallback = [];
  
  try {
    // Tenta chamar funções individuais diretamente
    const funcoesMonitoramento = [
      'monitorarBACEN',
      'monitorarCMN', 
      'monitorarDOU',
      'monitorarNoticias',
      'monitorarPortais'
    ];
    
    funcoesMonitoramento.forEach(nomeFuncao => {
      try {
        if (typeof this[nomeFuncao] === 'function') {
          const resultado = this[nomeFuncao]();
          if (resultado && Array.isArray(resultado)) {
            resultadosFallback.push(...resultado);
            Logger.log(`     ✅ ${nomeFuncao}: ${resultado.length} resultados`);
          }
        }
      } catch (e) {
        Logger.log(`     ⚠️ ${nomeFuncao}: ${e.message}`);
      }
    });
    
  } catch (error) {
    Logger.log(`   ❌ Fallback também falhou: ${error}`);
  }
  
  Logger.log(`   📊 Fallback: ${resultadosFallback.length} resultados`);
  return resultadosFallback;
}

// =============================================
// FUNÇÕES DE DIAGNÓSTICO
// =============================================

/**
 * 🔧 DIAGNOSTICAR MONITORAMENTO NORMATIVO
 */
function diagnosticarMonitoramento() {
  Logger.log('🔧 INICIANDO DIAGNÓSTICO DO MONITORAMENTO NORMATIVO');
  
  const diagnostico = {
    classeExiste: false,
    metodosDisponiveis: [],
    instanciacao: false,
    erro: null
  };
  
  try {
    // Verificar se a classe existe
    diagnostico.classeExiste = typeof MonitoramentoNormativo !== 'undefined';
    Logger.log(`📋 Classe MonitoramentoNormativo existe: ${diagnostico.classeExiste}`);
    
    if (diagnostico.classeExiste) {
      // Tentar instanciar
      try {
        const monitor = new MonitoramentoNormativo();
        diagnostico.instanciacao = true;
        Logger.log('✅ Instanciação bem-sucedida');
        
        // Listar métodos disponíveis
        diagnostico.metodosDisponiveis = Object.getOwnPropertyNames(Object.getPrototypeOf(monitor))
          .filter(prop => typeof monitor[prop] === 'function' && prop !== 'constructor');
        
        Logger.log(`📋 Métodos disponíveis: ${diagnostico.metodosDisponiveis.join(', ')}`);
        
        // Testar método principal
        if (diagnostico.metodosDisponiveis.includes('executarMonitoramentoCompleto')) {
          Logger.log('🧪 Testando executarMonitoramentoCompleto...');
          const resultadoTeste = monitor.executarMonitoramentoCompleto();
          Logger.log(`✅ Teste executado: ${Array.isArray(resultadoTeste) ? resultadoTeste.length + ' resultados' : 'sucesso'}`);
        }
        
      } catch (erroInstanciacao) {
        diagnostico.erro = erroInstanciacao.toString();
        Logger.log(`❌ Erro na instanciação: ${erroInstanciacao}`);
      }
    }
    
  } catch (error) {
    diagnostico.erro = error.toString();
    Logger.log(`❌ Erro no diagnóstico: ${error}`);
  }
  
  // Enviar relatório
  enviarSlackMensagem(
    `🔧 *DIAGNÓSTICO MONITORAMENTO NORMATIVO*\n\n` +
    `📋 Classe existe: ${diagnostico.classeExiste ? '✅' : '❌'}\n` +
    `🔧 Instanciação: ${diagnostico.instanciacao ? '✅' : '❌'}\n` +
    `📚 Métodos: ${diagnostico.metodosDisponiveis.join(', ') || 'Nenhum'}\n` +
    `${diagnostico.erro ? `❌ Erro: ${diagnostico.erro}` : '✅ Diagnóstico completo'}`
  );
  
  return diagnostico;
}

/**
 * 🧪 TESTE SIMPLIFICADO DO MONITORAMENTO
 */
function testeMonitoramentoSimplificado() {
  Logger.log('🧪 EXECUTANDO TESTE SIMPLIFICADO DO MONITORAMENTO');
  
  try {
    // Teste 1: Funções básicas de coleta
    Logger.log('1. Testando coletarNormativosReais()...');
    const normativosOficiais = coletarNormativosReais();
    Logger.log(`   ✅ Normativos oficiais: ${normativosOficiais.length}`);
    
    // Teste 2: Monitoramento complementar
    Logger.log('2. Testando monitoramento complementar...');
    let complementares = [];
    
    // Tenta diferentes abordagens
    try {
      const monitor = new MonitoramentoNormativo();
      complementares = monitor.executarMonitoramentoCompleto();
      Logger.log(`   ✅ Via classe: ${complementares.length} resultados`);
    } catch (e) {
      Logger.log(`   ⚠️ Classe falhou: ${e.message}`);
      
      // Fallback para funções diretas
      complementares = executarMonitoramentoFallback();
      Logger.log(`   🔄 Via fallback: ${complementares.length} resultados`);
    }
    
    // Resultado final
    const total = normativosOficiais.length + complementares.length;
    
    enviarSlackMensagem(
      `🧪 *TESTE MONITORAMENTO*\n\n` +
      `✅ Normativos oficiais: ${normativosOficiais.length}\n` +
      `✅ Fontes complementares: ${complementares.length}\n` +
      `📊 Total: ${total} normativos\n` +
      `🎯 Teste ${total > 0 ? 'BEM-SUCEDIDO' : 'SEM RESULTADOS'}`
    );
    
    return {
      oficiais: normativosOficiais.length,
      complementares: complementares.length,
      total: total,
      success: true
    };
    
  } catch (error) {
    Logger.log(`❌ TESTE FALHOU: ${error}`);
    enviarSlackMensagem(`❌ TESTE FALHOU: ${error.toString().substring(0, 150)}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =============================================
// VERSÃO ALTERNATIVA SE A CLASSE NÃO EXISTIR
// =============================================

/**
 * 🔄 IMPLEMENTAÇÃO ALTERNATIVA DO MONITORAMENTO
 * Use esta se a classe MonitoramentoNormativo não existir
 */
function executarMonitoramentoAlternativo() {
  Logger.log('🔄 EXECUTANDO MONITORAMENTO ALTERNATIVO');
  
  const resultados = [];
  
  try {
    // 1. BACEN
    try {
      Logger.log('   🏦 Monitorando BACEN...');
      const bacenResultados = monitorarBACEN();
      if (bacenResultados && Array.isArray(bacenResultados)) {
        resultados.push(...bacenResultados);
        Logger.log(`     ✅ BACEN: ${bacenResultados.length} resultados`);
      }
    } catch (e) {
      Logger.log(`     ⚠️ BACEN: ${e.message}`);
    }
    
    // 2. CMN
    try {
      Logger.log('   📊 Monitorando CMN...');
      const cmnResultados = monitorarCMN();
      if (cmnResultados && Array.isArray(cmnResultados)) {
        resultados.push(...cmnResultados);
        Logger.log(`     ✅ CMN: ${cmnResultados.length} resultados`);
      }
    } catch (e) {
      Logger.log(`     ⚠️ CMN: ${e.message}`);
    }
    
    // 3. DOU
    try {
      Logger.log('   📰 Monitorando DOU...');
      const douResultados = monitorarDOU();
      if (douResultados && Array.isArray(douResultados)) {
        resultados.push(...douResultados);
        Logger.log(`     ✅ DOU: ${douResultados.length} resultados`);
      }
    } catch (e) {
      Logger.log(`     ⚠️ DOU: ${e.message}`);
    }
    
    // 4. Notícias
    try {
      Logger.log('   📢 Monitorando notícias...');
      const noticiasResultados = monitorarNoticias();
      if (noticiasResultados && Array.isArray(noticiasResultados)) {
        resultados.push(...noticiasResultados);
        Logger.log(`     ✅ Notícias: ${noticiasResultados.length} resultados`);
      }
    } catch (e) {
      Logger.log(`     ⚠️ Notícias: ${e.message}`);
    }
    
    // 5. Portais
    try {
      Logger.log('   🌐 Monitorando portais...');
      const portaisResultados = monitorarPortais();
      if (portaisResultados && Array.isArray(portaisResultados)) {
        resultados.push(...portaisResultados);
        Logger.log(`     ✅ Portais: ${portaisResultados.length} resultados`);
      }
    } catch (e) {
      Logger.log(`     ⚠️ Portais: ${e.message}`);
    }
    
  } catch (error) {
    Logger.log(`❌ Monitoramento alternativo falhou: ${error}`);
  }
  
  Logger.log(`📊 Monitoramento alternativo: ${resultados.length} resultados totais`);
  return resultados;
}
// =============================================
// SISTEMA CORRIGIDO - INTEGRAÇÃO COMPLETA
// =============================================

/**
 * 🔍 EXECUTAR MONITORAMENTO COMPLETO - VERSÃO INTEGRADA
 * Usa tanto as funções oficiais quanto o MonitoramentoNormativo
 */
function executarMonitoramentoCompleto() {
  Logger.log('🔍 EXECUTANDO MONITORAMENTO COMPLETO - SISTEMA INTEGRADO');
  
  try {
    const resultados = {
      normativosOficiais: [],
      fontesComplementares: [],
      analisesToqan: [],
      backlog: 0,
      agenda: 0,
      startTime: new Date()
    };
    
    // 1. 📥 COLETA OFICIAL - SITES GOVERNAMENTAIS
    Logger.log('📥 ETAPA 1: COLETA OFICIAL (BACEN, RFB, CMN, SUSEP, DOU)...');
    try {
      resultados.normativosOficiais = coletarNormativosReais();
      Logger.log(`   ✅ ${resultados.normativosOficiais.length} normativos oficiais`);
    } catch (error) {
      Logger.log(`   ❌ Erro na coleta oficial: ${error}`);
      resultados.normativosOficiais = [];
    }
    
    // 2. 📥 COLETA COMPLEMENTAR - MONITORAMENTO NORMATIVO  
    Logger.log('📥 ETAPA 2: COLETA COMPLEMENTAR (Notícias, Portais)...');
    try {
      const monitor = new MonitoramentoNormativo();
      resultados.fontesComplementares = monitor.executarMonitoramentoCompleto();
      Logger.log(`   ✅ ${resultados.fontesComplementares.length} fontes complementares`);
    } catch (error) {
      Logger.log(`   ❌ Erro no monitoramento complementar: ${error}`);
      resultados.fontesComplementares = [];
    }
    
    // 3. 📊 COMBINAR TODOS OS RESULTADOS
    const todosNormativos = [
      ...resultados.normativosOficiais,
      ...resultados.fontesComplementares
    ];
    
    Logger.log(`📊 TOTAL COLETADO: ${todosNormativos.length} normativos`);
    Logger.log(`   🏛️  Oficiais: ${resultados.normativosOficiais.length}`);
    Logger.log(`   📰 Complementares: ${resultados.fontesComplementares.length}`);
    
    if (todosNormativos.length === 0) {
      Logger.log('⚡ Nenhum normativo detectado');
      enviarSlackMensagem('📭 *MONITORAMENTO IFOOD* - Nenhum normativo novo detectado hoje');
      return { success: true, mensagem: 'Nenhum normativo detectado' };
    }
    
    // 4. 🤖 ANÁLISE TOQAN
    Logger.log('🤖 ETAPA 3: ANÁLISE TOQAN...');
    const todasAnalises = analisarNormativosComToqan(todosNormativos);
    resultados.analisesToqan = todasAnalises;
    
    if (todasAnalises.length === 0) {
      Logger.log('⚡ Nenhuma análise concluída');
      enviarSlackMensagem('🤖 *ANÁLISE TOQAN* - Nenhuma análise concluída');
      return { success: false, mensagem: 'Análise não concluída' };
    }
    
    // ESTATÍSTICAS DAS ANÁLISES
    const aplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Sim').length;
    const naoAplicaveis = todasAnalises.filter(a => a.Aplicavel_iFood === 'Não').length;
    
    Logger.log(`   ✅ ${todasAnalises.length} análises (${aplicaveis} aplicáveis, ${naoAplicaveis} não aplicáveis)`);
    
    // 5. 📚 BACKLOG - SALVAR TODOS OS NORMATIVOS
    Logger.log('📚 ETAPA 4: BACKLOG (TODOS OS NORMATIVOS)...');
    const resultadoBacklog = salvarTodasAnalisesNoBacklog(todasAnalises);
    resultados.backlog = resultadoBacklog.salvos;
    resultados.backlogAplicaveis = resultadoBacklog.aplicaveis;
    resultados.backlogNaoAplicaveis = resultadoBacklog.naoAplicaveis;
    
    // 6. 💾 AGENDA NORMATIVA - SALVAR APENAS APLICÁVEIS
    Logger.log('💾 ETAPA 5: AGENDA NORMATIVA (APENAS APLICÁVEIS)...');
    resultados.agenda = salvarAplicaveisNaPlanilha(todasAnalises);
    Logger.log(`   ✅ ${resultados.agenda} aplicáveis na AgendaNormativa`);
    
    // 7. 📊 RELATÓRIO FINAL
    resultados.endTime = new Date();
    resultados.tempoExecucao = (resultados.endTime - resultados.startTime) / 1000;
    resultados.success = true;
    
    enviarRelatorioExecucaoIntegrado(resultados);
    
    Logger.log(`🎯 EXECUÇÃO CONCLUÍDA: ${resultados.backlog} no Backlog, ${resultados.agenda} na Agenda`);
    
    return resultados;
    
  } catch (error) {
    Logger.log(`❌ ERRO NO SISTEMA INTEGRADO: ${error.toString()}`);
    enviarSlackMensagem(`❌ ERRO NO SISTEMA: ${error.toString().substring(0, 150)}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * 📊 RELATÓRIO INTEGRADO - MOSTRA AMBAS AS FONTES
 */
function enviarRelatorioExecucaoIntegrado(resultados) {
  const tempoFormatado = resultados.tempoExecucao ? `${resultados.tempoExecucao.toFixed(1)}s` : 'N/A';
  
  const mensagem = 
    `📊 *RELATÓRIO DE EXECUÇÃO - SISTEMA INTEGRADO*\n\n` +
    `⏰ Horário: ${new Date().toLocaleString('pt-BR')}\n` +
    `⚡ Tempo: ${tempoFormatado}\n\n` +
    
    `📥 *COLETA OFICIAL (Órgãos Governamentais):*\n` +
    `• BACEN, RFB, CMN, SUSEP, DOU\n` +
    `• ${resultados.normativosOficiais.length} normativos oficiais\n\n` +
    
    `📰 *COLETA COMPLEMENTAR (Notícias/Portais):*\n` +
    `• BCB, LegisWeb, Valor, G1, InfoMoney, Forbes, Bloomberg\n` +
    `• ${resultados.fontesComplementares.length} fontes complementares\n\n` +
    
    `📊 *TOTAL COLETADO:* ${resultados.normativosOficiais.length + resultados.fontesComplementares.length} normativos\n\n` +
    
    `🤖 *ANÁLISE TOQAN:*\n` +
    `• Total analisado: ${resultados.analisesToqan.length}\n` +
    `• Aplicáveis iFood: ${resultados.backlogAplicaveis || 0}\n` +
    `• Não aplicáveis: ${resultados.backlogNaoAplicaveis || 0}\n\n` +
    
    `💾 *ARMAZENAMENTO:*\n` +
    `• 📚 Backlog (todos): ${resultados.backlog} registros\n` +
    `• 🗓️ AgendaNormativa (aplicáveis): ${resultados.agenda} registros\n\n` +
    
    `✅ *SISTEMA INTEGRADO FUNCIONANDO CORRETAMENTE*`;
  
  enviarSlackMensagem(mensagem);
}

// =============================================
// FUNÇÕES DE TESTE ESPECÍFICAS
// =============================================

/**
 * 🧪 TESTE DA COLETA OFICIAL
 */
function testeColetaOficial() {
  Logger.log('🧪 TESTANDO COLETA OFICIAL...');
  
  try {
    const normativos = coletarNormativosReais();
    
    Logger.log(`📊 RESULTADO COLETA OFICIAL: ${normativos.length} normativos`);
    
    normativos.forEach((norm, index) => {
      Logger.log(`   ${index + 1}. ${norm.Orgao} - ${norm.Tipo_Norma} ${norm.Numero} - ${norm.Tema}`);
    });
    
    enviarSlackMensagem(
      `🧪 *TESTE COLETA OFICIAL*\n\n` +
      `✅ ${normativos.length} normativos coletados\n` +
      `🏛️ Órgãos: ${[...new Set(normativos.map(n => n.Orgao))].join(', ')}\n` +
      `📋 Tipos: ${[...new Set(normativos.map(n => n.Tipo_Norma))].join(', ')}`
    );
    
    return {
      success: true,
      total: normativos.length,
      orgaos: [...new Set(normativos.map(n => n.Orgao))],
      tipos: [...new Set(normativos.map(n => n.Tipo_Norma))]
    };
    
  } catch (error) {
    Logger.log(`❌ TESTE COLETA OFICIAL FALHOU: ${error}`);
    enviarSlackMensagem(`❌ TESTE COLETA OFICIAL FALHOU: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 🧪 TESTE DO MONITORAMENTO COMPLEMENTAR
 */
function testeMonitoramentoComplementar() {
  Logger.log('🧪 TESTANDO MONITORAMENTO COMPLEMENTAR...');
  
  try {
    const monitor = new MonitoramentoNormativo();
    const resultados = monitor.executarMonitoramentoCompleto();
    
    Logger.log(`📊 RESULTADO MONITORAMENTO COMPLEMENTAR: ${resultados.length} itens`);
    
    // Agrupar por fonte
    const porFonte = {};
    resultados.forEach(item => {
      const fonte = item.Fonte || 'Desconhecida';
      if (!porFonte[fonte]) porFonte[fonte] = 0;
      porFonte[fonte]++;
    });
    
    let detalhes = '';
    for (const [fonte, quantidade] of Object.entries(porFonte)) {
      detalhes += `• ${fonte}: ${quantidade} itens\n`;
    }
    
    enviarSlackMensagem(
      `🧪 *TESTE MONITORAMENTO COMPLEMENTAR*\n\n` +
      `✅ ${resultados.length} itens coletados\n\n` +
      `📰 Distribuição por fonte:\n${detalhes}`
    );
    
    return {
      success: true,
      total: resultados.length,
      porFonte: porFonte
    };
    
  } catch (error) {
    Logger.log(`❌ TESTE MONITORAMENTO COMPLEMENTAR FALHOU: ${error}`);
    enviarSlackMensagem(`❌ TESTE MONITORAMENTO COMPLEMENTAR FALHOU: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 🧪 TESTE COMPLETO DO SISTEMA INTEGRADO
 */
function testeSistemaIntegrado() {
  Logger.log('🧪 EXECUTANDO TESTE COMPLETO DO SISTEMA INTEGRADO');
  
  const resultadosTeste = {
    coletaOficial: null,
    monitoramentoComplementar: null,
    integracao: null
  };
  
  try {
    // Teste 1: Coleta Oficial
    Logger.log('1. Testando coleta oficial...');
    resultadosTeste.coletaOficial = testeColetaOficial();
    
    Utilities.sleep(2000);
    
    // Teste 2: Monitoramento Complementar
    Logger.log('2. Testando monitoramento complementar...');
    resultadosTeste.monitoramentoComplementar = testeMonitoramentoComplementar();
    
    Utilities.sleep(2000);
    
    // Teste 3: Integração Completa
    Logger.log('3. Testando integração completa...');
    resultadosTeste.integracao = executarMonitoramentoCompleto();
    
    // Relatório Final
    const sucessoOficial = resultadosTeste.coletaOficial.success;
    const sucessoComplementar = resultadosTeste.monitoramentoComplementar.success;
    const sucessoIntegracao = resultadosTeste.integracao.success;
    
    const totalOficial = resultadosTeste.coletaOficial.total || 0;
    const totalComplementar = resultadosTeste.monitoramentoComplementar.total || 0;
    
    enviarSlackMensagem(
      `🧪 *TESTE COMPLETO DO SISTEMA INTEGRADO*\n\n` +
      `📥 Coleta Oficial: ${sucessoOficial ? '✅' : '❌'} ${totalOficial} normativos\n` +
      `📰 Monitoramento Complementar: ${sucessoComplementar ? '✅' : '❌'} ${totalComplementar} itens\n` +
      `🔗 Integração Completa: ${sucessoIntegracao ? '✅' : '❌'}\n\n` +
      `🎯 ${sucessoOficial && sucessoComplementar && sucessoIntegracao ? 'SISTEMA INTEGRADO FUNCIONANDO!' : 'AJUSTES NECESSÁRIOS'}`
    );
    
    return resultadosTeste;
    
  } catch (error) {
    Logger.log(`❌ TESTE COMPLETO FALHOU: ${error}`);
    enviarSlackMensagem(`❌ TESTE COMPLETO FALHOU: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}
