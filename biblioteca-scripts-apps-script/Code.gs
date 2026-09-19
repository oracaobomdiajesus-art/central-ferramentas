// Biblioteca de Scripts de Venda — GMN
//
// Ferramenta de uso pessoal (não é pra mandar link pra cliente): lista,
// filtra, edita e apaga scripts de venda organizados por nicho de público
// (rádio, outdoor, associação comercial, etc.), lendo e gravando direto
// numa planilha do Google Sheets.
//
// Como publicar:
// 1. Crie um projeto novo em script.google.com (Apps Script) — pode ser o
//    mesmo projeto do formulário de onboarding ou um novo, tanto faz.
// 2. Cole este arquivo como "Code.gs" e o "Index.html" como um arquivo HTML
//    chamado exatamente "Index".
// 3. No editor, selecione a função "configurar" e clique em Executar (rode
//    só uma vez) — autorize o acesso ao Google Sheets da sua conta.
// 4. Veja em "Execuções" (ou View > Logs) o link da planilha criada.
// 5. Clique em "Implantar" > "Nova implantação" > tipo "App da Web".
//    Executar como: Eu (sua conta). Quem pode acessar: **Só eu** — essa
//    página é só sua, diferente do formulário de onboarding.
// 6. Acesse o link gerado sempre que quiser consultar ou editar seus
//    scripts — também dá pra editar direto na planilha, os dois caminhos
//    ficam sincronizados.

const PLANILHA_PROPS_KEY = "BIBLIOTECA_PLANILHA_ID";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Biblioteca de Scripts — GMN")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Rode manualmente UMA VEZ (Executar > configurar) antes de publicar o App
// da Web: cria a planilha e guarda o ID dela pra reaproveitar sempre.
function configurar() {
  const propriedades = PropertiesService.getScriptProperties();
  let planilhaId = propriedades.getProperty(PLANILHA_PROPS_KEY);
  if (!planilhaId) {
    const planilha = SpreadsheetApp.create("Biblioteca de Scripts — GMN");
    const aba = planilha.getActiveSheet();
    aba.setName("Scripts");
    aba.appendRow(["ID", "Nicho", "Tipo de Script", "Texto do Script", "Última Atualização"]);
    aba.setFrozenRows(1);
    planilhaId = planilha.getId();
    propriedades.setProperty(PLANILHA_PROPS_KEY, planilhaId);
  }
  Logger.log("Planilha: https://docs.google.com/spreadsheets/d/" + planilhaId);
}

function getAba_() {
  const propriedades = PropertiesService.getScriptProperties();
  const planilhaId = propriedades.getProperty(PLANILHA_PROPS_KEY);
  if (!planilhaId) {
    throw new Error('Biblioteca ainda não configurada — rode a função "configurar" no Apps Script primeiro.');
  }
  return SpreadsheetApp.openById(planilhaId).getSheetByName("Scripts");
}

// Devolve todos os scripts cadastrados, pra montar a lista/filtro na página.
function listarScripts() {
  const aba = getAba_();
  const valores = aba.getDataRange().getValues();
  const linhas = valores.slice(1); // pula o cabeçalho
  return linhas
    .filter(function (linha) { return linha[0]; })
    .map(function (linha) {
      return {
        id: linha[0],
        nicho: linha[1],
        tipo: linha[2],
        texto: linha[3],
        atualizadoEm: linha[4] instanceof Date ? linha[4].toISOString() : String(linha[4] || ""),
      };
    })
    .sort(function (a, b) { return (b.atualizadoEm || "").localeCompare(a.atualizadoEm || ""); });
}

// Cria um script novo (sem "id") ou atualiza um existente (com "id") —
// usado tanto pelo formulário de "novo" quanto pelo de "editar" na página.
function salvarScript(dados) {
  const aba = getAba_();
  const agora = new Date();

  if (dados.id) {
    const valores = aba.getDataRange().getValues();
    for (let i = 1; i < valores.length; i++) {
      if (String(valores[i][0]) === String(dados.id)) {
        aba.getRange(i + 1, 2, 1, 4).setValues([[dados.nicho, dados.tipo, dados.texto, agora]]);
        return { ok: true, id: dados.id };
      }
    }
    throw new Error("Script não encontrado pra atualizar — pode ter sido apagado por outro acesso.");
  }

  const novoId = Utilities.getUuid();
  aba.appendRow([novoId, dados.nicho, dados.tipo, dados.texto, agora]);
  return { ok: true, id: novoId };
}

function excluirScript(id) {
  const aba = getAba_();
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) {
      aba.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error("Script não encontrado pra excluir — pode já ter sido apagado.");
}
