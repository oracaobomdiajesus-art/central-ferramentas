// Formulário de Onboarding de Clientes — GMN
//
// Como publicar:
// 1. Crie um projeto novo em script.google.com (Apps Script).
// 2. Cole este arquivo como "Code.gs" e "Index.html" e "GerarLink.html" como
//    dois arquivos HTML separados, com esses nomes exatos.
// 3. No editor, selecione a função "configurar" no menu de funções e clique
//    em Executar (rode só uma vez). Na primeira vez ele vai pedir autorização
//    de acesso ao Google Sheets e Drive da sua própria conta — autorize.
// 4. Veja em "Execuções" (ou View > Logs) os links da planilha e da pasta de
//    fotos que foram criadas — salve esses links pra você.
// 5. Clique em "Implantar" > "Nova implantação" > tipo "App da Web".
//    Executar como: Eu (sua conta). Quem pode acessar: Qualquer pessoa.
// 6. Copie o link do App da Web gerado, e acrescente "?painel=1" no final —
//    esse é o SEU link, de uso interno: marca os serviços que o cliente
//    contratou e ele gera o link certo pra você mandar pro cliente. O link
//    sem "?painel=1" é o formulário puro (só use direto se quiser que o
//    próprio cliente escolha os serviços).

const SHEET_PROPS_KEY = "PLANILHA_ID";
const PASTA_PROPS_KEY = "PASTA_ID";

function doGet(e) {
  const painel = Boolean(e && e.parameter && e.parameter.painel === "1");
  const pagina = painel ? "GerarLink" : "Index";
  return HtmlService.createHtmlOutputFromFile(pagina)
    .setTitle(painel ? "Gerar Link — GMN" : "Formulário de Onboarding — GMN")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Rode esta função manualmente UMA VEZ (Executar > configurar) antes de
// publicar o App da Web: cria a planilha de respostas e a pasta de fotos no
// seu Google Drive, e guarda os IDs delas pra reaproveitar em toda submissão
// futura (sem criar uma planilha nova a cada resposta).
function configurar() {
  const propriedades = PropertiesService.getScriptProperties();

  let planilhaId = propriedades.getProperty(SHEET_PROPS_KEY);
  if (!planilhaId) {
    const planilha = SpreadsheetApp.create("Onboarding de Clientes — GMN");
    const aba = planilha.getActiveSheet();
    aba.setName("Respostas");
    aba.appendRow([
      "Data/Hora", "Serviços Contratados",
      "Nome do Negócio", "Nome do Responsável", "Telefone/WhatsApp", "E-mail", "Endereço",
      "Instagram Atual", "Facebook Atual",
      "Categoria (GBP)", "Descrição (GBP)", "Horário de Funcionamento", "Já tem GBP?", "Link GBP Atual", "Fotos do Negócio (Drive)",
      "Nome Fantasia (Site)", "Quem Somos", "Produtos/Serviços", "Cor Preferida", "Fotos do Site (Drive)", "Chave Pix", "WhatsApp de Vendas",
      "Bio Curta (Biosite)", "Foto de Perfil (Drive)", "Links do Biosite",
    ]);
    aba.setFrozenRows(1);
    planilhaId = planilha.getId();
    propriedades.setProperty(SHEET_PROPS_KEY, planilhaId);
  }

  let pastaId = propriedades.getProperty(PASTA_PROPS_KEY);
  if (!pastaId) {
    const pasta = DriveApp.createFolder("Onboarding de Clientes — Fotos");
    pastaId = pasta.getId();
    propriedades.setProperty(PASTA_PROPS_KEY, pastaId);
  }

  Logger.log("Planilha: https://docs.google.com/spreadsheets/d/" + planilhaId);
  Logger.log("Pasta de fotos: https://drive.google.com/drive/folders/" + pastaId);
}

// Salva um grupo de fotos (já convertidas em base64 pelo navegador) numa
// subpasta própria da submissão, e devolve o link da subpasta — fica mais
// fácil de organizar (e revisitar depois) do que espalhar arquivo solto.
function salvarFotos_(nomeNegocio, fotos) {
  if (!fotos || !fotos.length) return "";

  const propriedades = PropertiesService.getScriptProperties();
  const pastaId = propriedades.getProperty(PASTA_PROPS_KEY);
  const pastaRaiz = DriveApp.getFolderById(pastaId);

  const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const nomePasta = (nomeNegocio || "sem-nome") + " - " + carimbo;
  const subPasta = pastaRaiz.createFolder(nomePasta);

  fotos.forEach(function (foto) {
    const bytes = Utilities.base64Decode(foto.base64);
    const blob = Utilities.newBlob(bytes, foto.mimeType, foto.nome);
    subPasta.createFile(blob);
  });

  return subPasta.getUrl();
}

// Chamada pelo formulário (via google.script.run) quando o cliente envia a
// resposta. Grava tudo numa linha nova da planilha "Respostas".
function salvarResposta(dados) {
  const propriedades = PropertiesService.getScriptProperties();
  const planilhaId = propriedades.getProperty(SHEET_PROPS_KEY);
  if (!planilhaId) {
    throw new Error('Formulário ainda não configurado — rode a função "configurar" no Apps Script primeiro.');
  }
  const planilha = SpreadsheetApp.openById(planilhaId).getSheetByName("Respostas");

  const linkFotosNegocio = salvarFotos_(dados.nomeNegocio, dados.fotosNegocio);
  const linkFotosSite = salvarFotos_(dados.nomeNegocio, dados.fotosSite);
  const linkFotoPerfil = salvarFotos_(dados.nomeNegocio, dados.fotoPerfil);

  const linksBiosite = (dados.linksBiosite || [])
    .filter(function (l) { return l.titulo || l.url; })
    .map(function (l) { return l.titulo + ": " + l.url; })
    .join(" | ");

  planilha.appendRow([
    new Date(),
    (dados.servicos || []).join(", "),
    dados.nomeNegocio || "", dados.nomeResponsavel || "", dados.telefone || "", dados.email || "", dados.endereco || "",
    dados.instagramAtual || "", dados.facebookAtual || "",
    dados.categoriaGbp || "", dados.descricaoGbp || "", dados.horarioFuncionamento || "", dados.jaTemGbp || "", dados.linkGbpAtual || "", linkFotosNegocio,
    dados.nomeFantasiaSite || "", dados.quemSomos || "", dados.produtosServicos || "", dados.corPreferida || "", linkFotosSite, dados.chavePix || "", dados.whatsappVendas || "",
    dados.bioBiosite || "", linkFotoPerfil, linksBiosite,
  ]);

  return { ok: true };
}
