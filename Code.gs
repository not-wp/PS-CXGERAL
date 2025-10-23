/************************************************************
 * Web App: Receitas + Atestado + Encaminhamento (PDF)
 * Lê "Mapeamento" por ID de planilha e gera apenas PDF no Drive
 ************************************************************/

// === PREENCHA com a planilha onde está a aba "Mapeamento" ===
const PLANILHA_MAPEAMENTO_ID = '1Xe_4GaMJgUaSzdxc1IQqz2asZhVCz_Is-3XFhHjpAdA';

// === CONSTANTES (ajuste aos títulos do seu formulário/fluxo) ===
const PERGUNTA_NOME      = 'Nome do paciente';
const PERGUNTA_RECEITAS  = 'Receita desejada';

const MAP_SHEET_NAME     = 'Mapeamento';
const PASTA_DESTINO_ID   = '1gJJ5EK5hHX97By-SZlVHtDZuz1f8Cs3A'; // pasta onde salvar PDFs

// Quebras de página (conforme seu último script salvo)
const INSERIR_QUEBRA_PAGINA_ENTRE_RECEITAS   = true;
const ADICIONAR_ASSINATURA                   = false;

// Atestado (opcional)
const PERGUNTA_ATESTADO   = 'Atestado?';            // checkbox Sim/Não
const PERGUNTA_DIAS_AT    = 'Dias de afastamento';  // número
const PERGUNTA_CID        = 'CID (opcional)';       // texto
const ROTULO_MAP_ATESTADO = 'Atestado';             // nome na coluna OpcaoForms
const INSERIR_QUEBRA_PAGINA_ANTES_ATESTADO   = false;

// Encaminhamento Ambulatorial (opcional)
const PERGUNTA_AMB_ESPECIALIDADE = 'Especialidade Ambulatório'; // texto livre
const PERGUNTA_AMB_SEMANAS       = 'Semanas até o retorno';     // número
const ROTULO_MAP_AMB             = 'Encaminhamento Ambulatorial';
const INSERIR_QUEBRA_PAGINA_ANTES_AMB = false;

/** ---------------- Web App endpoints ---------------- **/

/** UI do Web App (carrega index.html) */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Emissão de Documentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lista opções de receitas a partir do Mapeamento (exclui rótulos especiais) */
function listarReceitas() {
  const map = loadMapping_();
  const excluir = new Set([ROTULO_MAP_ATESTADO, ROTULO_MAP_AMB]);
  return Object.keys(map).filter(k => !excluir.has(k)).sort();
}

/**
 * Endpoint principal: recebe payload do front e retorna { pdfUrl, nomeArquivo }
 * payload: {
 *   nome, receitas[], atestadoSim, diasAtestado, cid,
 *   ambEspecialidade, ambSemanas
 * }
 */
function emitirDocumentos(payload) {
  const nomePaciente     = (payload?.nome || '').trim();
  const receitasSel      = Array.isArray(payload?.receitas) ? payload.receitas : [];
  const atestadoMarcado  = !!payload?.atestadoSim;
  const diasAtestado     = Math.max(1, parseInt(payload?.diasAtestado || '1', 10) || 1);
  const cidOpcional      = (payload?.cid || '').trim();

  const ambEspecialidade = (payload?.ambEspecialidade || '').trim();
  const temAmb           = ambEspecialidade.length > 0;

  // semanas pode vir '' (vazio) ou número; aceita 0
  const ambSemanasRaw    = (payload?.ambSemanas ?? '').toString().trim();
  const ambSemanasNum    = ambSemanasRaw === '' ? NaN : parseInt(ambSemanasRaw, 10);

  if (!nomePaciente) throw new Error('Informe o nome do paciente.');
  if (!receitasSel.length && !atestadoMarcado && !temAmb) {
    throw new Error('Selecione ao menos um documento: receita(s), atestado ou encaminhamento.');
  }

  // ==== VALIDAÇÃO REFORÇADA DO AMBULATÓRIO ====
  // Se tem especialidade preenchida, "Semanas" é obrigatório e deve ser número inteiro >= 0 (0 = na vaga)
  if (temAmb) {
    const semanasEhInteiroNaoNegativo =
      ambSemanasRaw !== '' &&
      /^-?\d+$/.test(ambSemanasRaw) &&           // formato numérico inteiro
      Number.isFinite(ambSemanasNum) &&
      ambSemanasNum >= 0;

    if (!semanasEhInteiroNaoNegativo) {
      throw new Error('Informe as semanas do retorno (use 0 se for na vaga).');
    }
  }
  // =============================================

  // Carrega mapeamento
  const mapa = loadMapping_();
  if (Object.keys(mapa).length === 0) {
    throw new Error(`A aba "${MAP_SHEET_NAME}" está vazia ou ausente.`);
  }

  // Cria Documento temporário
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  const docFinal = DocumentApp.create(`Receituário - ${nomePaciente} - ${ts}`);
  const bodyFinal = docFinal.getBody();

  // Helper: colar conteúdo de um modelo pelo DocID
  function appendFromModel_(docId) {
    const bodyModelo = DocumentApp.openById(docId).getBody();
    const n = bodyModelo.getNumChildren();
    for (let i = 0; i < n; i++) appendChild_(bodyFinal, bodyModelo.getChild(i));
    bodyFinal.appendParagraph('');
  }

  // Receitas
  receitasSel.forEach((opcao, idx) => {
    const map = mapa[opcao];
    if (!map) { Logger.log(`Sem mapeamento: "${opcao}"`); return; }
    if (INSERIR_QUEBRA_PAGINA_ENTRE_RECEITAS && idx > 0) bodyFinal.appendPageBreak();
    appendFromModel_(map.id);
  });

  // Atestado
  if (atestadoMarcado) {
    if (INSERIR_QUEBRA_PAGINA_ANTES_ATESTADO) bodyFinal.appendPageBreak();
    const mAt = mapa[ROTULO_MAP_ATESTADO];
    if (mAt) appendFromModel_(mAt.id);
  }

  // Encaminhamento
  if (temAmb) {
    if (INSERIR_QUEBRA_PAGINA_ANTES_AMB) bodyFinal.appendPageBreak();
    const mAmb = mapa[ROTULO_MAP_AMB];
    if (mAmb) appendFromModel_(mAmb.id);
  }

  // Substituições comuns
  const hoje = new Date();
  const dataHoje = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  bodyFinal.replaceText('<<\\s*NOME\\s*>>', nomePaciente);
  bodyFinal.replaceText('<<\\s*DATA\\s*>>', dataHoje);

  // Placeholders do Atestado
  if (atestadoMarcado) {
    // Hoje conta como 1º dia ⇒ fim = hoje + (dias - 1)
    const fim = new Date(hoje.getTime()); fim.setDate(fim.getDate() + (diasAtestado - 1));
    const dataFim = Utilities.formatDate(fim, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    bodyFinal.replaceText('<<\\s*DIAS_ATESTADO\\s*>>', String(diasAtestado));
    bodyFinal.replaceText('<<\\s*DATA_FIM_ATESTADO\\s*>>', dataFim);
    bodyFinal.replaceText('<<\\s*CID\\s*>>', cidOpcional || '');
  } else {
    bodyFinal.replaceText('<<\\s*DIAS_ATESTADO\\s*>>', '');
    bodyFinal.replaceText('<<\\s*DATA_FIM_ATESTADO\\s*>>', '');
    bodyFinal.replaceText('<<\\s*CID\\s*>>', '');
  }

  // Placeholders do Encaminhamento (pluralização e "na vaga.")
  if (temAmb) {
    bodyFinal.replaceText('<<\\s*ESPECIALIDADE\\s*>>', ambEspecialidade);

    let semanasTexto = '';
    if (Number.isFinite(ambSemanasNum)) {
      if (ambSemanasNum === 0) {
        semanasTexto = 'na vaga.';              // ajuste: tire o ponto se preferir no modelo
      } else if (ambSemanasNum === 1) {
        semanasTexto = 'em 1 semana';
      } else if (ambSemanasNum > 1) {
        semanasTexto = 'em ' + ambSemanasNum + ' semanas';
      }
    }
    bodyFinal.replaceText('<<\\s*RETORNO_SEMANAS\\s*>>', semanasTexto);
  } else {
    bodyFinal.replaceText('<<\\s*ESPECIALIDADE\\s*>>', '');
    bodyFinal.replaceText('<<\\s*RETORNO_SEMANAS\\s*>>', '');
  }

  if (ADICIONAR_ASSINATURA) {
    bodyFinal.appendParagraph('Assinatura: ______________________');
    bodyFinal.appendParagraph('CRM: XXXXX');
  }

  docFinal.saveAndClose();

  // Gera PDF e move para pasta destino; apaga .gdoc
  const pastaDestino = DriveApp.getFolderById(PASTA_DESTINO_ID);
  const fileDoc = DriveApp.getFileById(docFinal.getId());
  const pdfBlob = fileDoc.getAs('application/pdf');
  const pdfFile = pastaDestino.createFile(pdfBlob).setName(docFinal.getName() + '.pdf');
  fileDoc.setTrashed(true);

  // Compartilha por link (opcional)
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { pdfUrl: pdfFile.getUrl(), nomeArquivo: pdfFile.getName() };
}

/** ---------------- Helpers reutilizados ---------------- **/

// Anexa um filho do modelo ao body final respeitando o tipo
function appendChild_(dstBody, childEl) {
  const type = childEl.getType();
  switch (type) {
    case DocumentApp.ElementType.PARAGRAPH:
      dstBody.appendParagraph(childEl.asParagraph().copy().asParagraph());
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      dstBody.appendListItem(childEl.asListItem().copy().asListItem());
      break;
    case DocumentApp.ElementType.TABLE:
      dstBody.appendTable(childEl.asTable().copy().asTable());
      break;
    case DocumentApp.ElementType.PAGE_BREAK:
      dstBody.appendPageBreak();
      break;
    case DocumentApp.ElementType.HORIZONTAL_RULE:
      dstBody.appendHorizontalRule();
      break;
    case DocumentApp.ElementType.TABLE_OF_CONTENTS:
      dstBody.appendTableOfContents(childEl.asTableOfContents().copy().asTableOfContents());
      break;
    default:
      let txt = '';
      try { txt = childEl.getText(); } catch (_) {}
      if (txt) dstBody.appendParagraph(txt);
  }
}

// Extrai DocID de um ID puro ou URL do Docs/Drive
function extractDocId_(val) {
  val = String(val || '').trim();
  if (!val) return '';
  let m = val.match(/\/d\/([a-zA-Z0-9_-]+)/); if (m && m[1]) return m[1];
  let m2 = val.match(/[?&]id=([a-zA-Z0-9_-]+)/); if (m2 && m2[1]) return m2[1];
  return val;
}

// Carrega mapeamento da aba "Mapeamento" na planilha informada por ID
function loadMapping_() {
  const ss = SpreadsheetApp.openById(PLANILHA_MAPEAMENTO_ID);
  const sh = ss.getSheetByName(MAP_SHEET_NAME);
  if (!sh) throw new Error(`Crie a aba "${MAP_SHEET_NAME}" com 3 colunas: OpcaoForms | DocID | TituloOpcionalNoDocumento`);
  const last = sh.getLastRow();
  if (last < 2) return {};
  const values = sh.getRange(2, 1, last - 1, 3).getValues();
  const map = {};
  values.forEach(r => {
    const opc = (r[0] || '').toString().trim();
    const raw = (r[1] || '').toString().trim();
    const id  = extractDocId_(raw);
    const tit = (r[2] || '').toString().trim();
    if (opc && id) map[opc] = { id: id, titulo: (tit || opc) };
  });
  return map;
}
