const ID_PLANILHA = "1jXghWB8UgKcm6U72-doguM4auaDSQhvN7xt2QOTnc0w";
const ABA_ANALISE = "análise Dados";


// parte adicional para hospedar
function doGet(e) {
  const acao = e.parameter.acao;
  let resultado;

  try {
    if (acao === "definirMesAnoDashboard") {
      resultado = definirMesAnoDashboard(e.parameter.ano, e.parameter.mes);
    } 
    else if (acao === "definirMesAnoDefeitos") {
      resultado = definirMesAnoDefeitos(e.parameter.ano, e.parameter.mes);
    }
    else if (acao === "buscarPorData") {
      resultado = buscarPorData(e.parameter.dia, e.parameter.mes, e.parameter.ano);
    }
    else if (acao === "obterTotaisMes") {
      resultado = obterTotaisMes(e.parameter.mes, e.parameter.ano);
    }
    else if (acao === "obterTotalSemanaPorDia") {
      resultado = obterTotalSemanaPorDia(e.parameter.dia, e.parameter.mes, e.parameter.ano);
    }
    else if (acao === "obterResumoDefeitosCompleto") {
      resultado = obterResumoDefeitosCompleto();
    }
    else if (acao === "obterDadosGraficos") {
      resultado = obterDadosGraficos(e.parameter.mes, e.parameter.ano);
    }
    else if (acao === "obterListaProblemas") {
      resultado = obterListaProblemas(e.parameter.dia, e.parameter.mes, e.parameter.ano);
    }

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "erro": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}



function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Dashboard de Análise");
}

/* =========================================
   DASHBOARD – DEFINE L1 (ANO) E L2 (MÊS)
========================================= */
function definirMesAnoDashboard(ano, mesNumero) {

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const meses = [
    "janeiro","fevereiro","março","abril",
    "maio","junho","julho","agosto",
    "setembro","outubro","novembro","dezembro"
  ];

  sh.getRange("L1").setValue(String(ano));
  sh.getRange("L2").setValue(meses[Number(mesNumero)-1]);

  SpreadsheetApp.flush();
}

/* =========================================
   DEFEITOS + CARROS – DEFINE J18/J19
========================================= */
function definirMesAnoDefeitos(ano, mesNumero) {

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const meses = [
    "janeiro","fevereiro","março","abril",
    "maio","junho","julho","agosto",
    "setembro","outubro","novembro","dezembro"
  ];

  sh.getRange("J18").setValue(String(ano));
  sh.getRange("J19").setValue(meses[Number(mesNumero)-1]);

  SpreadsheetApp.flush();
  Utilities.sleep(300);

  return true;
}

/* =========================================
   BUSCA FICHAS POR DIA
========================================= */
function buscarPorData(dia, mes, ano) {

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const dados = sh.getRange("A12:G35").getDisplayValues();
  let resultado = [];

  dados.forEach(l => {

    if (!l[0]) return;

    const [d,m,a] = l[0].split("/").map(Number);

    if (
      a === Number(ano) &&
      m === Number(mes) &&
      (!dia || d === Number(dia))
    ){
      resultado.push([l[0], l[1], l[6]]);
    }
  });

  return resultado;
}

/* =========================================
   TOTAL MÊS + MÉDIA
========================================= */
function obterTotaisMes(mes, ano) {

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const dados = sh.getRange("A12:B35").getDisplayValues();
  let total = 0;

  dados.forEach(l => {

    if (!l[0]) return;

    const [d,m,a] = l[0].split("/").map(Number);

    if (m === Number(mes) && a === Number(ano)) {
      total += Number(l[1]) || 0;
    }
  });

  return {
    totalMes: total,
    mediaSemanal: sh.getRange("F2").getValue()
  };
}

/* =========================================
   TOTAL DA SEMANA (BASEADO NO DIA)
========================================= */
function obterTotalSemanaPorDia(dia, mes, ano) {

  if (!dia) return { totalSemana: "" };

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const dados = sh.getRange("A12:G35").getDisplayValues();

  let semanaAlvo = null;
  let totalSemana = 0;

  dados.forEach(l => {

    if (!l[0]) return;

    const [d,m,a] = l[0].split("/").map(Number);

    if (
      d === Number(dia) &&
      m === Number(mes) &&
      a === Number(ano)
    ){
      semanaAlvo = l[6];
    }
  });

  if (!semanaAlvo) return { totalSemana: "" };

  dados.forEach(l=>{
    if(l[6] === semanaAlvo){
      totalSemana += Number(l[1]) || 0;
    }
  });

  return { totalSemana };
}

function obterResumoDefeitosCompleto() {

  const sh = SpreadsheetApp
    .openById(ID_PLANILHA)
    .getSheetByName(ABA_ANALISE);

  const defeitos = sh.getRange("J23:L30")
    .getValues()
    .filter(l => l[0] && Number(l[1]) > 0);

  const carros = sh.getRange("N24:P30")
    .getValues()
    .filter(l => l[0] && Number(l[1]) > 0);

  const garagens = sh.getRange("N34:P37")
    .getValues()
    .filter(l => l[0] && Number(l[1]) > 0);

  const motoristas = sh.getRange("J35:L41")
    .getValues()
    .filter(l => l[0] && Number(l[2]) > 0);

  const mediaSemanal = sh.getRange("F2").getValue();  

  return {
    defeitos,
    carros,
    garagens,
    motoristas,
    mediaSemanal
  };
}

function obterDadosGraficos(mes, ano){

  const sh = SpreadsheetApp.openById(ID_PLANILHA)
                           .getSheetByName(ABA_ANALISE);

  const meses = [
    "janeiro","fevereiro","março","abril",
    "maio","junho","julho","agosto",
    "setembro","outubro","novembro","dezembro"
  ];

  sh.getRange("J18").setValue(String(ano));
  sh.getRange("J19").setValue(meses[Number(mes)-1]);

  SpreadsheetApp.flush();

  const defeitos = sh.getRange("J24:L30").getValues().filter(l=>l[0]);
  const carros = sh.getRange("N24:P30").getValues().filter(l=>l[0]);
  const garagens = sh.getRange("N35:P37").getValues().filter(l=>l[0]);

  return {
    defeitos: defeitos.map(l => ({
      nome: l[0],
      total: Number(l[1]),
      percentual: Number(l[2])
    })),
    carros: carros.map(l => ({
      nome: l[0],
      total: Number(l[1]),
      percentual: Number(l[2])
    })),
    garagens: garagens.map(l => ({
      nome: l[0],
      total: Number(l[1]),
      percentual: Number(l[2])
    }))
  };
}


/* =========================================
   BUSCA DETALHAMENTO DE PROBLEMAS (A50:B200)
========================================= */
function obterListaProblemas(dia, mes, ano) {
  const sh = SpreadsheetApp.openById(ID_PLANILHA).getSheetByName(ABA_ANALISE);

  const meses = [
    "janeiro","fevereiro","março","abril",
    "maio","junho","julho","agosto",
    "setembro","outubro","novembro","dezembro"
  ];

  // Atualiza os filtros específicos da lista detalhada
  sh.getRange("L1").setValue(String(ano));
  sh.getRange("L2").setValue(meses[Number(mes) - 1]);
  sh.getRange("I45").setValue(dia ? Number(dia) : "");

  SpreadsheetApp.flush();
  Utilities.sleep(500); // Tempo para o Sheets processar a lista

  // Captura o intervalo A50:B200 e filtra linhas vazias
  const dados = sh.getRange("A50:B200").getValues().filter(l => l[0] !== "");

  return dados;
}
