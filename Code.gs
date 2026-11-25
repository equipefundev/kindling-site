/**
 * KINDLING — Lista de espera (Google Planilhas + notificação por e-mail)
 *
 * CONFIGURAÇÃO:
 * 1. Abra a planilha de destino > Extensões > Apps Script.
 * 2. Cole este código no editor.
 * 3. Defina EMAIL_DESTINO com o e-mail que receberá as notificações.
 * 4. Execute configurarPlanilha() uma vez e autorize o acesso à planilha/e-mail.
 * 5. Implante como Aplicativo da Web:
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 * 6. Copie a URL terminada em /exec para o index.html.
 */

const NOME_ABA = "Lista de espera";
const EMAIL_DESTINO = "SEU_EMAIL@DOMINIO.COM";

function configurarPlanilha() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  if (!planilha) {
    throw new Error("Abra este projeto Apps Script a partir da planilha de destino.");
  }

  let aba = planilha.getSheetByName(NOME_ABA);
  if (!aba) aba = planilha.insertSheet(NOME_ABA);

  const cabecalhos = [
    "Data de inscrição",
    "Nome",
    "E-mail",
    "Escola/Empresa",
    "WhatsApp",
    "Tipo de uso",
    "Quantidade de usuários",
    "Consentimento",
    "Página de origem"
  ];

  if (aba.getLastRow() === 0) {
    aba.appendRow(cabecalhos);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight("bold")
      .setBackground("#2b2119")
      .setFontColor("#ffffff");
    aba.autoResizeColumns(1, cabecalhos.length);
  }
}

/**
 * Recebe o formulário por POST, grava primeiro na planilha e tenta enviar
 * a notificação por e-mail. Retorna uma página que envia o resultado ao site
 * por postMessage, evitando o falso positivo do fetch no-cors.
 */
function doPost(e) {
  let resultado;

  try {
    const dados = (e && e.parameter) ? e.parameter : {};
    const requestId = limpar(dados.requestId);

    // Campo invisível: se estiver preenchido, provavelmente é um robô.
    if (limpar(dados.website)) {
      return paginaResposta("success", "Inscrição recebida.", true, requestId);
    }

    const nome = limpar(dados.name);
    const email = limpar(dados.email);
    const organizacao = limpar(dados.organization);
    const whatsapp = limpar(dados.whatsapp);
    const tipo = limpar(dados.type);
    const tamanho = limpar(dados.size);
    const consentimento = dados.consent === "on";

    if (!nome || !email || !organizacao || !whatsapp || !tipo || !consentimento) {
      return paginaResposta("error", "Confira os campos obrigatórios e tente novamente.", false, requestId);
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return paginaResposta("error", "O endereço de e-mail informado não é válido.", false, requestId);
    }

    // Filtro simples de envios automatizados excessivamente rápidos.
    const enviadoEm = Number(dados.submittedAt || 0);
    if (enviadoEm && (Date.now() - enviadoEm) < 1800) {
      return paginaResposta("error", "Não foi possível validar o envio. Tente novamente.", false, requestId);
    }

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    if (!planilha) throw new Error("Planilha vinculada não encontrada.");

    let aba = planilha.getSheetByName(NOME_ABA);
    if (!aba) {
      configurarPlanilha();
      aba = planilha.getSheetByName(NOME_ABA);
    }

    const agora = new Date();
    aba.appendRow([
      agora,
      protegerFormula(nome),
      protegerFormula(email),
      protegerFormula(organizacao),
      protegerFormula(whatsapp),
      protegerFormula(tipo),
      protegerFormula(tamanho),
      "Sim",
      protegerFormula(limpar(dados.page))
    ]);

    // A gravação é a etapa principal. Falha no e-mail não apaga a inscrição.
    let emailEnviado = true;
    let avisoEmail = "";

    if (EMAIL_DESTINO && EMAIL_DESTINO !== "SEU_EMAIL@DOMINIO.COM") {
      try {
        const assunto = "Nova inscrição na lista de espera do Kindling";
        const corpo =
          "Uma nova pessoa demonstrou interesse em testar o Kindling.\n\n" +
          "Nome: " + nome + "\n" +
          "E-mail: " + email + "\n" +
          "Escola/Empresa: " + organizacao + "\n" +
          "WhatsApp: " + whatsapp + "\n" +
          "Tipo de uso: " + tipo + "\n" +
          "Quantidade de usuários: " + (tamanho || "Não informado") + "\n" +
          "Data: " + Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + "\n\n" +
          "Acesse a planilha para acompanhar os contatos.";

        MailApp.sendEmail({
          to: EMAIL_DESTINO,
          subject: assunto,
          body: corpo,
          replyTo: email,
          name: "Lista de espera Kindling"
        });
      } catch (emailErro) {
        emailEnviado = false;
        console.error("Inscrição gravada, mas o e-mail falhou: " + emailErro);
        avisoEmail = " Sua inscrição foi salva, mas a notificação por e-mail não foi enviada.";
      }
    } else {
      emailEnviado = false;
      avisoEmail = " A planilha recebeu sua inscrição, mas o e-mail de notificação ainda não foi configurado.";
    }

    resultado = {
      status: "success",
      message: "Inscrição registrada com sucesso! A equipe da Fundev entrará em contato." + avisoEmail,
      emailSent: emailEnviado
    };
    return paginaResposta(resultado.status, resultado.message, resultado.emailSent, requestId);

  } catch (erro) {
    console.error(erro);
    return paginaResposta("error", "Ocorreu um erro ao registrar sua inscrição. Tente novamente em alguns instantes.", false, requestId);
  }
}

/** Página de retorno para o iframe invisível no site. */
function paginaResposta(status, mensagem, emailEnviado, requestId) {
  const payload = JSON.stringify({
    kind: "KINDLING_WAITLIST_RESULT",
    status: status,
    message: mensagem,
    emailSent: Boolean(emailEnviado),
    requestId: String(requestId || "")
  }).replace(/</g, "\\u003c");

  const html =
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>' +
    'window.top.postMessage(' + payload + ', "*");' +
    '</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Limita tamanho e remove espaços externos. */
function limpar(valor) {
  return String(valor || "").trim().slice(0, 500);
}

/** Evita que texto iniciado por caracteres especiais seja interpretado como fórmula. */
function protegerFormula(valor) {
  const texto = String(valor || "");
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}
